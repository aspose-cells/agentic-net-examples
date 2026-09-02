// Title: Generate a separate CSV file for each worksheet in an Aspose.Cells workbook using a custom delimiter and UTF-8 encoding (C#)
// AI Prompts: Write C# code that loops through all worksheets in a Workbook and saves each one as an individual CSV file with a user‑defined separator and UTF‑8 encoding using Aspose.Cells TxtSaveOptions. | Create a reusable method that accepts a workbook path, output folder, delimiter character, and text encoding, then exports every sheet to separate CSV files named after the worksheets. | Adjust the CSV export logic to support any delimiter (comma, semicolon, tab) and any encoding (UTF‑8, UTF‑16) while exporting only the active sheet.
// Common Searches: Aspose.Cells C# export each worksheet to its own CSV file with semicolon delimiter | how to set custom separator and UTF-8 encoding when saving Excel sheets as CSV using Aspose.Cells | C# loop through workbook worksheets and save as separate CSV files with TxtSaveOptions | generate multiple CSV files from one Excel workbook using Aspose.Cells .NET API
// Tags: Aspose.Cells TxtSaveOptions CSV export per worksheet | custom delimiter CSV export Aspose.Cells | UTF-8 encoding CSV Aspose.Cells | export workbook sheets to individual CSV files C# | save active worksheet as CSV Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // The program creates a workbook with two worksheets, iterates over each sheet, sets it as active, configures TxtSaveOptions with a semicolon separator and UTF-8 encoding, and saves the active sheet as a CSV file named after the worksheet.
    class Program
    {
        static void Main()
        {
            // Create a workbook and add sample worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "First";
            workbook.Worksheets[0].Cells["A1"].PutValue("Name");
            workbook.Worksheets[0].Cells["B1"].PutValue("Age");
            workbook.Worksheets[0].Cells["A2"].PutValue("Alice");
            workbook.Worksheets[0].Cells["B2"].PutValue(30);

            // Add a second worksheet with different data
            Worksheet sheet2 = workbook.Worksheets.Add("Second");
            sheet2.Cells["A1"].PutValue("Product");
            sheet2.Cells["B1"].PutValue("Price");
            sheet2.Cells["A2"].PutValue("Apple");
            sheet2.Cells["B2"].PutValue(1.5);

            // Define a custom delimiter (e.g., semicolon) and UTF‑8 encoding
            char customDelimiter = ';';
            Encoding utf8 = Encoding.UTF8;

            // Export each worksheet to its own CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Set the current worksheet as active
                workbook.Worksheets.ActiveSheetIndex = i;

                // Configure TxtSaveOptions for CSV export
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = customDelimiter,   // custom delimiter
                    Encoding = utf8,               // UTF‑8 encoding
                    ExportAllSheets = false        // export only the active sheet
                };

                // Build output file name based on worksheet name
                string sheetName = workbook.Worksheets[i].Name;
                string outputPath = $"{sheetName}.csv";

                // Save the active worksheet as CSV using the configured options
                workbook.Save(outputPath, saveOptions);
            }

            Console.WriteLine("CSV files have been generated for each worksheet.");
        }
    }
}
