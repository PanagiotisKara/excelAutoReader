using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

[Route("api/excel")]
[ApiController]
public class ExcelController : ControllerBase
{
    [HttpPost("upload")]
    [Consumes("multipart/form-data")]
    public IActionResult UploadExcel(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { message = "Please upload a valid Excel file." });

        List<object> extractedData = new List<object>();

        using (var stream = new MemoryStream())
        {
            file.CopyTo(stream);
            stream.Position = 0;

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true 
                    }
                });

                DataTable table = result.Tables[0]; 

                Console.WriteLine("Column Names in Excel:");
                foreach (DataColumn col in table.Columns)
                {
                    Console.WriteLine(col.ColumnName);
                }
                //<-----------------------------------------------------------choosing excel column names-------------------------------------------------------------->
                foreach (DataRow row in table.Rows)
                {
                    string crmbatrid = table.Columns.Contains("ID Εξόδου") && row["ID Εξόδου"] != DBNull.Value ? row["ID Εξόδου"].ToString() : "N/A";
                    string preInvoicedocNo = table.Columns.Contains("Αρ. Προτιμολογίου") && row["Αρ. Προτιμολογίου"] != DBNull.Value ? row["Αρ. Προτιμολογίου"].ToString() : "N/A";
                    string BaTR_Amount = table.Columns.Contains("Πληρωτέο") && row["Πληρωτέο"] != DBNull.Value ? row["Πληρωτέο"].ToString() : "N/A";
                    string Notes = table.Columns.Contains("Comments") && row["Comments"] != DBNull.Value ? row["Comments"].ToString() : "N/A";

                    extractedData.Add(new
                    {
                        crmbatrid = crmbatrid,               
                        preInvoicedocNo = preInvoicedocNo,
                        BaTR_Amount = BaTR_Amount,
                        Notes = Notes
                    });
                }

            }
        }

        return Ok(new { data = extractedData });
    }
}
