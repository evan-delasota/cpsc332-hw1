# This script continues from Company.ps1, it assumes the CPSC332 database already exists.
# Uncomment the command below to create the CPSC332 database if needed.
# Invoke-SQLCmd -HostName "localhost" -Database "CPSC332" -InputFile ".\00. Company.sql" -ErrorAction Stop;

$OutputFile = ".\DepartmentStats.csv";
$SqlQuery = "
SELECT FirstCol.DepartmentName,
    SumOfSalary,
    AvgOfSalary,
    SumOfWorkingHours,
    TotalNumOfProjects
FROM
(
    SELECT DEPARTMENT.Dname AS DepartmentName, TotalNumOfProjects FROM
    (
        SELECT PROJECT.Dnum,
            Count(DISTINCT WORKS_ON.Pno) AS TotalNumOfProjects
            FROM   WORKS_ON
            INNER JOIN PROJECT 
            ON PROJECT.Pnumber = WORKS_ON.Pno
            GROUP  BY PROJECT.Dnum
    ) AS p
    INNER JOIN DEPARTMENT    
    ON DEPARTMENT.Dnumber = p.Dnum
) AS FirstCol
    
INNER JOIN 
(
    SELECT SecondCol.DepartmentName,
    ThirdCol.SumOfSalary,
    ThirdCol.AvgOfSalary,
    SecondCol.SumOfWorkingHours
    FROM
    (
        SELECT FirstQuery.Dname  AS DepartmentName,
        Sum(SecondQuery.HOURS) AS SumOfWorkingHours
        FROM DEPARTMENT AS FirstQuery
        FULL JOIN
        (
            SELECT * FROM   WORKS_ON
            FULL JOIN PROJECT 
            ON WORKS_ON.Pno = PROJECT.Pnumber
        ) AS SecondQuery 
        ON firstQuery.Dnumber = SecondQuery.Dnum
        GROUP  BY FirstQuery.Dname
    ) AS SecondCol
    
    INNER JOIN 
    (
        SELECT * FROM
        (
            SELECT (d.Dname) AS DepartmentName,
            Sum(e.salary) AS SumOfSalary,
            Avg(e.salary) AS AvgOfSalary
            FROM employee AS e
                    RIGHT OUTER JOIN
                    DEPARTMENT AS d
                                ON
                    d.Dnumber = e.Dno
            GROUP  BY d.Dname
        ) AS e
    ) AS ThirdCol 
    ON SecondCol.DepartmentName = ThirdCol.DepartmentName

) AS FourthCol
ON FirstCol.DepartmentName = FourthCol.DepartmentName";

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
$SqlConnection.ConnectionString = "Server = ETD; Database = CPSC332; Integrated Security = True";
$SqlConnection.Open();

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
$SqlCmd.Connection = $SqlConnection;
$SqlCmd.CommandText = $SqlQuery;

$SqlConnection.Close();
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
$SqlAdapter.SelectCommand = $SqlCmd;

$DataSet = New-Object System.Data.DataSet;
$SqlAdapter.Fill($DataSet);
$SqlConnection.Close();

$DataSet.Tables[0] | select "DepartmentName", "SumOfSalary", "AvgOfSalary", "SumOfWorkingHours", "TotalNumOfProjects" | Export-Csv $OutputFile;