// Title: Auto‑fit every row in an Excel worksheet with Aspose.Cells for .NET using AutoFitRows from the first to the last data row
// Common Searches: C# Aspose.Cells auto fit rows from first to last data row | How to use AutoFitRows to adjust all row heights in a .NET Excel workbook | Aspose.Cells AutoFitRows startRow 0 endRow MaxDataRow example | Auto‑fit rows in an existing Excel file using Aspose.Cells C# | Set row height automatically for all rows in a worksheet with Aspose.Cells
// Tags: auto-fit rows Aspose.Cells C# | AutoFitRows startRow endRow usage | MaxDataRow determination for row auto-fit | adjust row height Excel worksheet .NET | load or create workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file (or creates a new workbook if none exists), accesses the first worksheet, determines the last populated row using MaxDataRow, and calls AutoFitRows(0, lastRow) to automatically adjust the height of every row based on its content before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load an existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one empty sheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Find the index of the last row that contains data (zero‑based)
            int lastRow = sheet.Cells.MaxDataRow;

            // Auto‑fit rows only when there is at least one data row
            if (lastRow >= 0)
            {
                sheet.AutoFitRows(0, lastRow);
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
