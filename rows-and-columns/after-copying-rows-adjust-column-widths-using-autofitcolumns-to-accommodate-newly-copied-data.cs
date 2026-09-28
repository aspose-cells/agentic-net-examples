// Title: Copy rows and auto‑fit column widths in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that copies a block of rows to a new position and then calls Worksheet.AutoFitColumns to resize all columns. | Demonstrate how to use Cells.CopyRows followed by Worksheet.AutoFitColumns in Aspose.Cells to duplicate rows and adjust column widths. | Generate a self‑contained example that copies rows 2‑4 to row 6 and automatically fits column widths in an Excel file with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy rows to another location and autofit columns | How to use Cells.CopyRows and then AutoFitColumns in .NET | C# example for copying rows and adjusting column width with Aspose.Cells | AutoFitColumns after copying rows in Excel using Aspose.Cells for .NET | Resize Excel columns automatically after row duplication with Aspose.Cells
// Tags: cells.copierows method Aspose.Cells C# | worksheet.autofitcolumns usage Aspose.Cells | excel column width adjustment after row copy | duplicate rows in worksheet Aspose.Cells | auto‑fit columns programmatically Aspose.Cells | copy rows and resize columns .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, fills cells A1:B4 with sample data, copies rows 2‑4 to start at row 6 using Cells.CopyRows, invokes Worksheet.AutoFitColumns to automatically size the columns for the new data, and saves the result as CopyRowsAndAutoFitColumnsDemo.xlsx.
    public class CopyRowsAndAutoFitColumnsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source rows with sample data
            cells["A1"].PutValue("Header");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Row 1");
            cells["B2"].PutValue(12345);
            cells["A3"].PutValue("Row 2");
            cells["B3"].PutValue(67890);
            cells["A4"].PutValue("Row 3");
            cells["B4"].PutValue(11121);

            // Copy rows 1-3 (zero‑based indices 1 to 3) to start at row 5
            // sourceRowIndex = 1, destinationRowIndex = 5, rowNumber = 3
            cells.CopyRows(cells, 1, 5, 3);

            // After copying, adjust column widths to fit the new data
            sheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "CopyRowsAndAutoFitColumnsDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
