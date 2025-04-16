This project demonstrates how to build an ASP.NET MVC application that allows users to upload Excel files (.xlsx or .xls), parse the data using ExcelDataReader, and insert the data into a SQL Server database table. The uploaded data is also displayed on the browser for confirmation.

✅ Features
Upload Excel file through a simple web form

Parse Excel data using ExcelDataReader

Display uploaded data in an HTML table

Insert the data into SQL Server database (TestingUpload table)

Automatically detects and uses Excel headers (like Name, RegNo, Department)

Remove duplicate entries from the database using SQL

Clean, beginner-friendly code structure following MVC pattern

🛠 Technologies Used
ASP.NET MVC (.NET 6 / .NET Core)

C#

SQL Server 2022 Express

Entity Framework or ADO.NET (depending on version)

ExcelDataReader for reading Excel files

🧾 Table Structure
SQL Server Table: TestingUpload

sql
Copy
Edit
CREATE TABLE TestingUpload (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Name NVARCHAR(100),
    RegNo NVARCHAR(100),
    Department NVARCHAR(100)
);
📤 How It Works
User selects an Excel file via the web form

The app reads the file and extracts rows using ExcelDataReader

Data is inserted into SQL Server using parameterized queries

Inserted data is shown in a responsive HTML table

Optional: Duplicate rows can be removed via a SQL query

🗑 SQL to Remove Duplicate Rows
sql
Copy
Edit
WITH Duplicates AS (
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY Name, RegNo, Department
               ORDER BY Id
           ) AS rn
    FROM TestingUpload
)
DELETE FROM Duplicates
WHERE rn > 1;
