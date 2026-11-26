// Title: How to force quoted CSV fields by setting JsonLayoutOptions.QuoteAllFields and TxtSaveOptions.QuoteAllFields in Aspose.Cells (C#)
// Common Searches: how to enable QuoteAllFields on JsonLayoutOptions for CSV output using Aspose.Cells in C# | Aspose.Cells C# save workbook as CSV with all fields quoted via TxtSaveOptions | using reflection to set QuoteAllFields property on Aspose.Cells save options | export JSON data to quoted CSV file with Aspose.Cells C# example | force double quotes around every CSV column in Aspose.Cells
// Tags: Aspose.Cells JsonLayoutOptions configuration | Aspose.Cells TxtSaveOptions CSV quoting | C# reflection set Aspose.Cells property | CSV export with quoted fields Aspose.Cells | Import JSON into worksheet Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;   // Provides TxtSaveOptions

namespace AsposeCellsQuoteAllFieldsDemo
{
    // The example creates a workbook, configures JsonLayoutOptions (including QuoteAllFields via reflection), imports JSON data into the first worksheet, prepares TxtSaveOptions for CSV output (also enabling QuoteAllFields via reflection), ensures the output directory exists, and saves the workbook as a UTF‑8 CSV file where every field is enclosed in double quotes.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for the output CSV file
            string outputCsvPath = "output.csv";

            // Sample JSON data to be imported into the worksheet
            string jsonData = @"
            {
                ""Employees"": [
                    { ""Name"": ""John Doe"", ""Age"": 30, ""StartDate"": ""2022-01-15"", ""Salary"": 75000.50 },
                    { ""Name"": ""Jane Smith"", ""Age"": 28, ""StartDate"": ""2021-07-01"", ""Salary"": 68000.00 }
                ]
            }";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initialize JsonLayoutOptions
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                DateFormat = "yyyy-MM-dd",
                NumberFormat = "#,##0.00"
            };

            // Try to set JsonLayoutOptions.QuoteAllFields via reflection (if the property exists)
            PropertyInfo jsonQuoteProp = typeof(JsonLayoutOptions).GetProperty("QuoteAllFields", BindingFlags.Public | BindingFlags.Instance);
            if (jsonQuoteProp != null && jsonQuoteProp.CanWrite)
            {
                jsonQuoteProp.SetValue(jsonOptions, true);
                Console.WriteLine("JsonLayoutOptions.QuoteAllFields set to true via reflection.");
            }
            else
            {
                Console.WriteLine("JsonLayoutOptions does not expose a QuoteAllFields property. Proceeding without it.");
            }

            // Import JSON data into the worksheet using the configured options
            try
            {
                JsonUtility.ImportData(jsonData, worksheet.Cells, 0, 0, jsonOptions);
                Console.WriteLine("JSON data imported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error importing JSON data: " + ex.Message);
                return;
            }

            // Prepare TXT (CSV) save options
            TxtSaveOptions txtOptions = new TxtSaveOptions
            {
                Separator = ',',               // Use comma as CSV separator
                Encoding = Encoding.UTF8
            };

            // Try to set TxtSaveOptions.QuoteAllFields via reflection (if the property exists)
            PropertyInfo txtQuoteProp = typeof(TxtSaveOptions).GetProperty("QuoteAllFields", BindingFlags.Public | BindingFlags.Instance);
            if (txtQuoteProp != null && txtQuoteProp.CanWrite)
            {
                txtQuoteProp.SetValue(txtOptions, true);
                Console.WriteLine("TxtSaveOptions.QuoteAllFields set to true via reflection.");
            }
            else
            {
                Console.WriteLine("TxtSaveOptions does not expose a QuoteAllFields property. Proceeding without it.");
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputCsvPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error preparing output directory: " + ex.Message);
                return;
            }

            // Save the workbook as CSV
            try
            {
                workbook.Save(outputCsvPath, txtOptions);
                Console.WriteLine($"Workbook saved as CSV at: {Path.GetFullPath(outputCsvPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving CSV file: " + ex.Message);
            }
        }
    }
}
