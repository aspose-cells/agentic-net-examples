// Title: How to auto‑fit rows 15‑20 in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that fills rows 15‑20 with long text, enables text wrapping, and calls Worksheet.AutoFitRows to resize those rows. | Show how to apply a wrapped‑text style to a cell range, auto‑fit the row heights, and save the workbook as an .xlsx file with Aspose.Cells. | Create a console application that creates a workbook, writes sample data to rows 15‑20, auto‑fits the rows, and saves the file to a user‑specified folder.
// Common Searches: Aspose.Cells C# auto fit rows 15 to 20 example | Worksheet.AutoFitRows startRow endRow parameters usage | C# adjust Excel row height automatically with text wrap using Aspose.Cells | How to auto‑size specific rows in an Excel file with Aspose.Cells .NET | Save Aspose.Cells workbook to desktop folder C#
// Tags: Worksheet.AutoFitRows range | auto fit row height Aspose.Cells | text wrap style Aspose.Cells C# | save workbook as xlsx Aspose.Cells | populate rows with long text Aspose.Cells

using System;
using Aspose.Cells;

namespace AutoFitRowsRangeDemo
{
    // The example creates a new workbook, writes long text into rows 15‑20 of column A, enables text wrapping for those cells, calls Worksheet.AutoFitRows(14, 19) to automatically adjust the row heights, and saves the workbook as AutoFitRowsRangeDemo.xlsx on the desktop.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate rows 15‑20 (zero‑based indices 14‑19) with sample data
            for (int row = 14; row <= 19; row++)
            {
                // Add some long text to column A to demonstrate row height adjustment
                worksheet.Cells[row, 0].PutValue($"This is a sample text for row {row + 1} that is intentionally long to require auto‑fitting of the row height.");
                
                // Optionally enable text wrapping so the height changes
                Style style = worksheet.Cells[row, 0].GetStyle();
                style.IsTextWrapped = true;
                worksheet.Cells[row, 0].SetStyle(style);
            }

            // Auto‑fit rows 15‑20 (indices 14‑19)
            worksheet.AutoFitRows(14, 19);

            // Save the workbook to the desktop (adjust path as needed)
            string outputPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "AutoFitRowsRangeDemo.xlsx");

            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
