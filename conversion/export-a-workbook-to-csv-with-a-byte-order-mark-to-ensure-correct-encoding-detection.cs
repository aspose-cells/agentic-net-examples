// Title: Export Workbook to CSV with UTF‑8 BOM using Aspose.Cells for C#
// Description: Demonstrates how to create a workbook, populate cells, and save it as a CSV file that includes a UTF‑8 Byte Order Mark. The example uses TxtSaveOptions (SaveFormat.Csv), sets Encoding to UTF8, defines a comma separator, and optionally exports all worksheets.
// Keywords: Aspose.Cells | CSV export | UTF-8 BOM | C# | TxtSaveOptions | SaveFormat.Csv | ExportAllSheets | .NET encoding | Unicode CSV file
// Common Searches: Aspose.Cells C# export CSV with BOM | How to add UTF-8 Byte Order Mark to CSV in .NET | TxtSaveOptions Encoding UTF8 CSV Aspose | Save workbook as CSV with UTF-8 in C# | Export all worksheets to a single CSV file Aspose
// Developer Intent: Generate a CSV file from an Aspose.Cells workbook that contains a UTF‑8 Byte Order Mark for reliable encoding detection.
// Use Cases: Produce CSV reports that open correctly in Excel, Notepad, and other tools by embedding a UTF‑8 BOM. | Automate data pipelines where downstream systems require CSV files to be UTF‑8 compliant with a BOM. | Export multiple worksheets into one CSV while preserving Unicode characters.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as a CSV file using UTF‑8 encoding and a Byte Order Mark, specifying the separator and exporting all sheets. | Explain how TxtSaveOptions properties (Encoding, Separator, ExportAllSheets) affect CSV output and how to configure them to include a BOM in Aspose.Cells for .NET.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvBomDemo
{
    // Demonstrates how to create a workbook, populate cells, and save it as a CSV file that includes a UTF‑8 Byte Order Mark. The example uses TxtSaveOptions (SaveFormat.Csv), sets Encoding to UTF8, defines a comma separator, and optionally exports all worksheets.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure CSV save options with UTF-8 encoding (includes BOM)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8,   // Ensures a BOM is written
                Separator = ',',           // Standard comma separator
                ExportAllSheets = true     // Export all sheets (optional)
            };

            // Save the workbook as CSV with the specified options
            workbook.Save("output_with_bom.csv", csvOptions);

            Console.WriteLine("Workbook exported to CSV with UTF-8 BOM.");
        }
    }
}
