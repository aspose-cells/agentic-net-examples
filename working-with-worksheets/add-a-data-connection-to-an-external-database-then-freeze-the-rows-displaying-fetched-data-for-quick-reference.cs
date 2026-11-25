// Title: Create an Excel workbook, simulate a database data pull, and freeze the header row with Aspose.Cells for .NET (C#)
// Common Searches: how to simulate a database query and fill an Aspose.Cells worksheet in C# | Aspose.Cells C# freeze top row after importing data | populate Excel sheet from external source and apply FreezePanes using Aspose.Cells .NET | save workbook as XLSX after freezing header row with Aspose.Cells C# | example of creating workbook and freezing rows for quick reference in Aspose.Cells
// Tags: populate worksheet from database Aspose.Cells | freeze header row Aspose.Cells | save workbook as XLSX Aspose.Cells | create new workbook C# Aspose.Cells | simulate data connection Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataConnectionAndFreeze
{
    // The sample program creates a new Workbook, adds placeholder data that mimics a database fetch, freezes the first (header) row using FreezePanes, ensures the output directory exists, and saves the file as DatabaseDataWithFreeze.xlsx in XLSX format using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "DatabaseDataWithFreeze.xlsx";

            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "FetchedData";

            // -----------------------------------------------------------------
            // 2‑4. Add sample data (external‑connection APIs may be unavailable)
            // -----------------------------------------------------------------
            try
            {
                // Populate the worksheet with sample data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Sample Item 1");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Sample Item 2");
                Console.WriteLine("Sample data inserted.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inserting sample data: {ex.Message}");
            }

            // 5. Freeze the top row (header)
            try
            {
                // Freeze the first row (row index 1) and no columns
                sheet.FreezePanes(1, 0, 1, 0);
                Console.WriteLine("Top row frozen for quick reference.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error applying freeze panes: {ex.Message}");
            }

            // 6. Save the workbook
            try
            {
                // Ensure the directory for the output file exists
                string fullPath = Path.GetFullPath(outputPath);
                string directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
