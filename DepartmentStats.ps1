# This script continues from Company.ps1, it assumes the CPSC332 database already exists.
# Uncomment the command below to create the CPSC332 database if needed.
# Invoke-SQLCmd -HostName "localhost" -Database "CPSC332" -InputFile ".\00. Company.sql" -ErrorAction Stop;

$OutputFile = ".\DepartmentStats.csv";
$SqlQuery = "SELECT Dname AS DepartmentName, SUM(DISTINCT EMPLOYEE.Salary) AS SumOfSalary,
AVG(DISTINCT Salary) AS AvgOfSalary, SUM(ISNULL(WORKS_ON.Hours, 0)) AS SumOfWorkingHours, 
COUNT(DISTINCT Pno) AS TotalNumOfProjects

FROM EMPLOYEE
INNER JOIN WORKS_ON
ON EMPLOYEE.Ssn = WORKS_ON.Essn
INNER JOIN DEPARTMENT
ON Employee.Dno = DEPARTMENT.Dnumber
GROUP BY DEPARTMENT.Dname;";

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
$SqlConnection.ConnectionString = "Server = ETD; Database = CPSC332; Integrated Security = True";
$SqlConnection.Open();

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
$SqlCmd.Connection = $SqlConnection;

$SqlCmd.CommandText = $SqlQuery;

# while ($Data.Read()) {
#     Write-Host $Data["DepartmentName", "SumOfSalary", "AvgOfSalary", "SumOfWorkingHours", "TotalNumOfProjects"];
# };

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
$SqlAdapter.SelectCommand = $SqlCmd;

$DataSet = New-Object System.Data.DataSet;
$SqlAdapter.Fill($DataSet);
$SqlConnection.Close();

$DataSet.Tables[0] | select "DepartmentName", "SumOfSalary", "AvgOfSalary", "SumOfWorkingHours", "TotalNumOfProjects" | Export-Csv $OutputFile;