## Writeback Email Address to Aeries SIS for Staff
## The purpose of this script is to write back email addresses from AD
## to Aeries SIS records. This will ensure the records are kept in sync

## Instructions
## 1. Update Server and Database Names

$config = @{ 
                server = "SERVER NAME";
                database = "DATABASE NAME";
}

$connectionString = "Data Source=$($config.server);Initial Catalog=$($config.database);Integrated Security=SSPI;";

$Query = "SELECT ID, SID, EM,AEM FROM dbo.STF";

function GetDataFromSqlDatabase($Query, $ConnectionString) {
    try {
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
        $SqlConnection.ConnectionString = $ConnectionString;
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
        $SqlCmd.Connection = $SqlConnection;
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
        
        $SqlCmd.CommandText = $query;
        $SqlAdapter.SelectCommand = $SqlCmd;
        
        $DataSet = New-Object System.Data.DataSet;
        $SqlAdapter.Fill($DataSet) | out-null;
        return $DataSet.Tables[0] | Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors;
    }
    catch {
        Write-Error "Something went wrong while connecting to the SQL server";
        Write-Error $_.Exception.Message;
    }
    finally 
    {
        if($SqlConnection -and $SqlConnection.State -eq [System.Data.ConnectionState]::Open)
        {
            $SqlConnection.Close()
        }
    }
}

function UpdateDataInSqlDatabase($Query, $ConnectionString) {
    try {
        # Initialize connection and query information
        # Connect to the SQL server
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
        $SqlConnection.ConnectionString = $ConnectionString;
        $SqlConnection.open();
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
        $SqlCmd.Connection = $SqlConnection;
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
        
        $SqlCmd.CommandText = $query;
        $SqlAdapter.SelectCommand = $SqlCmd;
        $sqlcmd.ExecuteNonQuery();
        $SqlConnection.close();
    }
    catch {
        Write-Error "Something went wrong while connecting to the SQL server";
        Write-Error $_.Exception.Message;
    }
    finally 
    {
        if($SqlConnection -and $SqlConnection.State -eq [System.Data.ConnectionState]::Open)
        {
            $SqlConnection.Close()
        }
    }
}

#Get Staff
$aeriesData = GetDataFromSqlDatabase -Query $Query -ConnectionString $connectionString

#Get AD Users
$adUsers = Get-ADUser -LDAPFilter "(&(employeeID=*)(mail=*))" -Properties employeeID, mail

#PreProcess AD Users
$orderedAD = [ordered]@{};
foreach($i in $adUsers)
{
    $orderedAD[$i.EmployeeID] = $i.mail;
}

#Loop of Aeries Users, check and update
foreach($i in $aeriesData)
{
    try{
    
    $ADMail = $orderedAD["$($i.ID)"];
    if($ADMail -ne '' -and $ADMail -ne $Null)
    {
        if($i.EM -ne $ADMail)
        {
            $query = "UPDATE dbo.STF SET EM = '$($ADMail)'  WHERE ID = '$($i.ID)'"
            $result = UpdateDataInSqlDatabase -Query $Query -ConnectionString $connectionString
            $query = $null;
        }
    }
     
    }
    catch
    {
        Write-Error -Verbose $_;
    }
    
}