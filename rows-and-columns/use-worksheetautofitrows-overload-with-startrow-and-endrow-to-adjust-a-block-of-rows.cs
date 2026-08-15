// Title: Auto‑fit a range of rows in Aspose.Cells (C#) using Worksheet.AutoFitRows(startRow, endRow)
// Description: Shows how to create a workbook, fill rows with short, long, and wrapped text, and automatically adjust the height of rows 0‑4 with Worksheet.AutoFitRows(startRow, endRow) before saving the file.
// Keywords: Aspose.Cells | Worksheet.AutoFitRows | C# auto‑fit rows range | adjust Excel row height programmatically | partial row auto‑fit .NET | wrapped text row height | Excel row height optimization
// Common Searches: Aspose.Cells auto‑fit specific rows C# | Worksheet.AutoFitRows overload example | How to resize only certain rows in an Excel file using Aspose.Cells | AutoFitRows startRow endRow usage | Fit rows 1‑5 in Aspose.Cells .NET
// Developer Intent: Resize only rows 0‑4 to their optimal heights while leaving all other rows untouched.
// Use Cases: Fit header rows that contain wrapped or long text after data generation. | Adjust a dynamic block of rows added in a loop without altering preset row heights. | Prepare a report where only the title and summary sections need automatic height adjustment.
// AI Prompts: Generate C# code that uses Worksheet.AutoFitRows to resize rows 2‑8 after applying cell styles. | Provide an example combining Worksheet.AutoFitRows and Worksheet.AutoFitColumns for full worksheet formatting in Aspose.Cells. | Explain how to calculate startRow and endRow based on the longest cell content before calling AutoFitRows.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, fill rows with short, long, and wrapped text, and automatically adjust the height of rows 0‑4 with Worksheet.AutoFitRows(startRow, endRow) before saving the file.
    public class AutoFitRowsBlockDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data that will affect row heights
                for (int row = 0; row < 10; row++)
                {
                    // Long text in column A for odd rows to demonstrate auto‑fit
                    string text = (row % 2 == 0)
                        ? "Short text"
                        : "This is a considerably longer piece of text that should cause the row height to increase after AutoFitRows is applied.";
                    cells[row, 0].PutValue(text);

                    // Add a second column with wrapped text
                    cells[row, 1].PutValue("Wrapped\ntext\nexample");

                    // Enable text wrapping for the second column
                    Style wrapStyle = cells[row, 1].GetStyle();
                    wrapStyle.IsTextWrapped = true;
                    cells[row, 1].SetStyle(wrapStyle);
                }

                // Define the block of rows to auto‑fit (rows 0 through 4 inclusive)
                int startRow = 0;
                int endRow = 4;

                // Auto‑fit the specified range of rows
                worksheet.AutoFitRows(startRow, endRow);

                // Save the workbook to disk
                string outputPath = "AutoFitRowsBlockDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            AutoFitRowsBlockDemo.Run();
        }
    }
}
