using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using ExcelDataReader;
using WebApplication3_upload_excel_file_and_then_access_it_.Models;

public class HomeController : Controller
{
    private readonly string _connectionString = "Server=DESKTOP-2CTNAI7\\SQLEXPRESS;Database=db1;Trusted_Connection=True;TrustServerCertificate=True;";

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> Upload(IFormFile file)
    {
        var records = new List<TestingUpload>();

        try
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (file != null && file.Length > 0)
            {
                using (var stream = file.OpenReadStream())
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    var table = result.Tables[0];

                    using (SqlConnection conn = new SqlConnection(_connectionString))
                    {
                        await conn.OpenAsync();

                        foreach (DataRow row in table.Rows)
                        {
                            string name = row["Name"]?.ToString();
                            string regNo = row["RegNo"]?.ToString();
                            string dept = row["Department"]?.ToString();

                            records.Add(new TestingUpload
                            {
                                Name = name,
                                RegNo = regNo,
                                Department = dept
                            });

                            string query = "INSERT INTO TestingUpload (Name, RegNo, Department) VALUES (@Name, @RegNo, @Department)";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@Name", name ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@RegNo", regNo ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@Department", dept ?? (object)DBNull.Value);

                                await cmd.ExecuteNonQueryAsync();
                            }
                        }
                    }
                }

                ViewBag.Message = $"{records.Count} records inserted successfully!";
            }
            else
            {
                ViewBag.Message = "No file selected.";
            }
        }
        catch (Exception ex)
        {
            ViewBag.Error = "❌ Error: " + ex.Message;
        }

        ViewBag.Data = records;
        return View("Index");
    }
}
