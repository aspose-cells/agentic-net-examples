// Title: Automatically create new worksheets when importing a large CSV with Aspose.Cells TxtLoadOptions in C#
// AI Prompts: Write C# code that loads a CSV using Aspose.Cells and enables automatic sheet continuation when the row count exceeds Excel’s limit. | Show how to set a custom maximum rows per worksheet before a new sheet is added during CSV import with Aspose.Cells. | Demonstrate adding the header row to each generated worksheet while using Aspose.Cells to split a CSV across multiple sheets.
// Common Searches: Aspose.Cells C# import CSV and split into multiple worksheets automatically | how to handle Excel 1,048,576 row limit when loading large CSV with Aspose.Cells | C# TxtLoadOptions ExtendToNextSheet example for CSV files | auto create new sheet when CSV rows exceed limit using Aspose.Cells .NET | preserve header row on each worksheet during CSV to Excel conversion Aspose.Cells
// Tags: Aspose.Cells TxtLoadOptions ExtendToNextSheet | CSV to multi-sheet Excel import C# | automatic worksheet creation for large CSV | custom row limit per worksheet Aspose.Cells | repeat header row on each generated sheet

using System;
using System.IO;
using Aspose.Cells;

namespace AutoPopulateAcrossSheets
{
    // Loads a CSV with Aspose.Cells using TxtLoadOptions set to ExtendToNextSheet, automatically adds new worksheets when the row limit is reached, auto‑fits columns, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source CSV file.
                string inputCsvPath = @"C:\Data\LargeDataSet.csv";

                // Path where the resulting workbook will be saved.
                string outputXlsxPath = @"C:\Data\LargeDataSet.xlsx";

                // Ensure the input CSV exists; create a small sample if it does not.
                if (!File.Exists(inputCsvPath))
                {
                    Console.WriteLine($"Input file not found: {inputCsvPath}");
                    Directory.CreateDirectory(Path.GetDirectoryName(inputCsvPath));
                    File.WriteAllLines(inputCsvPath, new[]
                    {
                        "Id,Name,Value",
                        "1,Alpha,100",
                        "2,Beta,200",
                        "3,Gamma,300"
                    });
                    Console.WriteLine("Sample CSV file created.");
                }

                // Configure loading options to automatically continue data on a new worksheet.
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    ExtendToNextSheet = true
                };

                // Load the CSV file into a workbook using the configured options.
                Workbook workbook = new Workbook(inputCsvPath, loadOptions);

                // Auto‑fit columns for better readability.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.AutoFitColumns();
                }

                // Ensure the output directory exists.
                Directory.CreateDirectory(Path.GetDirectoryName(outputXlsxPath));

                // Save the workbook.
                workbook.Save(outputXlsxPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook created successfully with {workbook.Worksheets.Count} worksheets.");
                Console.WriteLine($"Saved to: {outputXlsxPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
