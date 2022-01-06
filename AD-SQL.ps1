
# main try block
Try
{

    # Import SqlServer Module
    if (Get-Module -Name sqlps) { Remove-Module sqlps }
    Import-Module -Name SqlServer

    # ~~~~~~~~~~~~~~ Active Directory ~~~~~~~~~~~~~~
    # creat organizational unit and populate with personnel records

    # creating organizational unit
    $ADRoot = "DC=consultingfirm,DC=com"
    $OUName = "finance"
    $OUDisplayName = "Finance"
    $ADPath = "OU=$($OUName),$($ADPath)"
    # if statement check if OU exists // if not doesn't run code and gives warning
    if (-Not([ADSI]::Exists("LDAP://$($ADPath)"))) {
        New-ADOrganizationalUnit -Path $ADRoot -Name $OUName -DisplayName $OUDisplayName -ProtectedFromAccidentalDeletion $false

        # read CSV file into a table
        $NewADUsers = Import-Csv -Path $PSScriptRoot\financePersonnel.csv
        $Path = "OU=finance,DC=consultingfirm,DC=com"

        # set up values for status reporting
        $numberNewUsers = $NewADUsers.Count
        $count = 1

        # iterate over each row in the table
        foreach ($ADUser in $NewADUsers)
        {
            # assign variable to column values
            $First = $ADUser.First_Name
            $Last = $ADUser.Last_Name
            $Name = $First + " " + $Last 
            $SamAcct = $ADUser.samAccount
            $Postal = $ADUser.PostalCode
            $Office = $ADUser.OfficePhone
            $Mobile = $ADUser.MobilePhone
            
            # create and show status
            $status = "Adding AD User: $($Name) ($($count) of $($numberNewUsers))"
            Write-Progress -Activity 'C916 Task 2 - Restore' -Status $status -PercentComplete (($count/$numberNewUsers) * 100)

            # user values to create each AD user
            New-ADUser `
            -GivenName $First `
            -Surname $Last `
            -Name $Name `
            -SamAccountName $SamAcct `
            -DisplayName $Name `
            -PostalCode $Postal `
            -MobilePhone $Mobile `
            -OfficePhone $Office `
            -Path $Path

            # increment counter
            $count++
        }

        Write-Host -ForegroundColor Cyan "Active Directory Tasks Complete"

    # end of if statement
    }
    else
    {
        Write-Host -ForegroundColor Red "$($OUName) Organizational Unit Already Exists"
    }

# ~~~~~~~~~~~~~~ SQL Server ~~~~~~~~~~~~~~

    # SQL server try block
    try
    {

        # string variable of SQL instance
        $sqlServerInstanceName = "SRV19-PRIMARY\SQLEXPRESS"

        # create object reference to the SQL server 
        $sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstanceName

        # name of the database
        $databaseName = "ClientDB"

        # create object reference to the database
        $databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName

        # call the create method on database object
        $databaseObject.Create()

        # execute T-SQL code against database
        Invoke-Sqlcmd -ServerInstance $sqlServerInstanceName -Database $databaseName -InputFile $PSScriptRoot\CreateTable_Client_A_Contacts.sql 

        # adding records from csv
        $tableName = "Client_A_Contacts"
        $Insert = "INSERT INTO [$($tableName)] (first_name, last_name, city, county, zip, officePHone, mobilePhone) "

        # structure contains new records
        $NewCustomerLeads = Import-Csv $PSScriptRoot\NewClientData.csv

        # loop formats VALUES portion of the INSERT INTO statement
        ForEach($NewLead in $NewCustomerLeads)
        {
            $Values = "VALUES ( `
            '$($NewLead.first_name)', `
            '$($NewLead.last_name)', `
            '$($NewLead.city)', `
            '$($NewLead.county)', `
            '$($NewLead.zip)', `
            '$($NewLead.officePhone)', `
            '$($NewLead.mobilePhone)')"

        $query = $Insert + $Values
        Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $query 
        }

        Write-Host -ForegroundColor Cyan "SQL Tasks Complete"

    # end of SQL try block
    }

    Catch
    {
        Write-Host -ForegroundColor Cyan "SQL Database Already Exists"
    }

# end of main try block
}

# catch system out of memory error
Catch [System.OutOfMemoryException]
{
    Write-Host -ForegroundColor Red "OutOfMemoryException Occured"
}

# catch all other errors
Catch
{
    Write-Host -ForegroundColor Red "Error: $($_.Exception.Message)"
}
