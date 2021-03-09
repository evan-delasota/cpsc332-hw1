
## Invoke-SQLCmd -HostName "localhost" -Database "CPSC332" -InputFile ".\00. Company.sql" -ErrorAction Stop;

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
$SqlConnection.ConnectionString = "Server = ETD; Database = CPSC332; Integrated Security = True";
$SqlConnection.Open();

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
$SqlCmd.Connection = $SqlConnection;
"`r`nDepartment Tables`r`n"

# Employee Table Output
"Employees"
$SqlCmd.CommandText = "SELECT * FROM dbo.EMPLOYEE;";

$Data = $SqlCmd.ExecuteReader();

while ($Data.Read()) {
    Write-Host $Data["Fname", "Minit", "Lname", "Ssn", "Bdate", "Address", "Sex", "Salary", "Super_ssn", "Dno"]
};
"`r`n"
$SqlConnection.Close();

# Department Table Output
"Departments"
$SqlConnection.Open();
$SqlCmd.CommandText = "SELECT * FROM dbo.DEPARTMENT;";

$Data = $SqlCmd.ExecuteReader();
while ($Data.Read()) {
    Write-Host $Data["Dname", "Dnumber", "Mgr_ssn", "Mgr_start_date"];
};
"`r`n"
$SqlConnection.Close();

# Department_Location Table Output 
"Department Locations"
$SqlConnection.Open();
$SqlCmd.CommandText = "SELECT * FROM dbo.DEPT_LOCATIONS;";

$Data = $SqlCmd.ExecuteReader();
while ($Data.Read()) {
    Write-Host $Data["Dnumber", "Dlocation"];
};
"`r`n"
$SqlConnection.Close();

# Project Table Output
"Projects"
$SqlConnection.Open();
$SqlCmd.CommandText = "SELECT * FROM dbo.PROJECT;";

$Data = $SqlCmd.ExecuteReader();
while ($Data.Read()) {
    Write-Host $Data["Pname", "Pnumber", "Plocation", "Dnum"];
};
"`r`n"
$SqlConnection.Close();

# Works_On Table Output
"Employee ID And Tasked Project"
$SqlConnection.Open();
$SqlCmd.CommandText = "SELECT * FROM dbo.WORKS_ON;";

$Data = $SqlCmd.ExecuteReader();
while ($Data.Read()) {
    Write-Host $Data["Essn", "Pno", "Hours"];
};
"`r`n"
$SqlConnection.Close();



# Dependent Table Output
"Dependents"
$SqlConnection.Open();
$SqlCmd.CommandText = "SELECT * FROM dbo.DEPENDENT;";

$Data = $SqlCmd.ExecuteReader();
while ($Data.Read()) {
    Write-Host $Data["Essn", "Dependent_name", "Sex", "Bdate", "Relationship"];
};
$SqlConnection.Close();
"`r`n`r`n"
