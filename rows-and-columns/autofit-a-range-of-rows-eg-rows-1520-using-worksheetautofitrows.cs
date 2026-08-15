// Title: Auto‑fit rows 15‑20 in an Aspose.Cells workbook with Worksheet.AutoFitRows (C#)
// Description: Creates a new Workbook, writes long wrapped text to rows 15‑20 (zero‑based 14‑19), enables text wrapping, calls Worksheet.AutoFitRows(14, 19) to resize the row heights, and saves the file as an XLSX document on the desktop.
// Keywords: Aspose.Cells | Worksheet.AutoFitRows | C# | .NET | auto fit rows | row height adjustment | text wrap in Excel | range of rows | Excel export | programmatic row sizing
// Common Searches: Aspose.Cells auto fit specific rows C# | Worksheet.AutoFitRows example with text wrapping | How to adjust row height for a range in Aspose.Cells | C# code to auto‑fit rows 15 to 20 in Excel | Save workbook after auto‑fitting rows Aspose.Cells
// Developer Intent: Automatically resize the height of rows 15‑20 based on their wrapped content using Aspose.Cells for .NET.
// Use Cases: Generating reports where certain rows contain wrapped text that must be auto‑sized before distribution. | Populating a predefined row block in a template and ensuring the rows fit the inserted data. | Exporting user‑entered data to Excel and dynamically adjusting selected row heights for optimal readability.
// AI Prompts: Write C# code that uses Worksheet.AutoFitRows to auto‑fit rows 10‑15 after enabling text wrapping. | Show how to accept a user‑defined row range, enable text wrap, auto‑fit those rows, and save the workbook with Aspose.Cells. | Explain the behavior of Worksheet.AutoFitRows with wrapped cells and how zero‑based row indexes affect the method call.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, writes long wrapped text to rows 15‑20 (zero‑based 14‑19), enables text wrapping, calls Worksheet.AutoFitRows(14, 19) to resize the row heights, and saves the file as an XLSX document on the desktop.
    public class AutoFitRowsRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate rows 15‑20 (zero‑based indexes 14‑19) with sample data
                for (int row = 14; row <= 19; row++)
                {
                    // Put a long text in column A to demonstrate row height adjustment
                    worksheet.Cells[row, 0].PutValue(
                        $"This is a long text for row {row + 1} that should cause the row height to increase after AutoFitRows is applied.");

                    // Enable text wrapping to see the effect more clearly
                    Style style = worksheet.Cells[row, 0].GetStyle();
                    style.IsTextWrapped = true;
                    worksheet.Cells[row, 0].SetStyle(style);
                }

                // Auto‑fit rows 15‑20 (indexes 14‑19)
                worksheet.AutoFitRows(14, 19);

                // Save the workbook to the desktop (adjust the path as needed)
                string outputPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "AutoFitRowsRangeDemo.xlsx");

                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutoFitRowsRangeDemo.Run();
        }
    }
}
