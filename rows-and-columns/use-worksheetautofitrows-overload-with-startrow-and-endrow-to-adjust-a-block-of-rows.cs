// Title: Auto‑fit a specific block of rows in an Aspose.Cells worksheet with Worksheet.AutoFitRows(startRow, endRow) – C# example
// AI Prompts: Write C# code that creates a workbook, fills several cells with long text, enables text wrapping, and then calls Worksheet.AutoFitRows with start and end indices to resize only those rows. | Show how to auto‑size rows 1‑5 in an Aspose.Cells sheet by using the AutoFitRows overload that accepts startRow and endRow, then save the workbook. | Provide a step‑by‑step C# snippet that demonstrates populating a column, applying wrapping style, invoking AutoFitRows for a row range, and exporting the file.
// Common Searches: Aspose.Cells C# how to auto‑fit rows 1 to 5 only | Worksheet.AutoFitRows(startRow, endRow) example for .NET | Resize a selected range of rows in an Aspose.Cells workbook | C# code to apply text wrapping then auto‑adjust row heights in Excel file
// Tags: Aspose.Cells auto‑fit rows block | Worksheet.AutoFitRows startRow endRow overload | C# text wrapping before row autofit | export Excel file after adjusting row heights | adjust row heights for long cell content Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, writes long text into cells A1‑A5, enables text wrapping, calls Worksheet.AutoFitRows(0, 4) to auto‑adjust the heights of rows 1‑5, and saves the result as AutoFitRowsBlockDemo.xlsx.
    public class AutoFitRowsBlockDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will affect row heights
            sheet.Cells["A1"].PutValue("This is a very long text that should cause the first row to expand after autofit.");
            sheet.Cells["A2"].PutValue("Short text");
            sheet.Cells["A3"].PutValue("Another long piece of text that will demonstrate row autofit for a range of rows.");
            sheet.Cells["A4"].PutValue("Medium length text");
            sheet.Cells["A5"].PutValue("Final row with a long text to test the autofit range.");

            // Enable text wrapping to see height changes more clearly
            for (int row = 0; row <= 4; row++)
            {
                Style style = sheet.Cells[row, 0].GetStyle();
                style.IsTextWrapped = true;
                sheet.Cells[row, 0].SetStyle(style);
            }

            // AutoFit rows from index 0 to 4 (i.e., rows 1 through 5)
            sheet.AutoFitRows(0, 4);

            // Save the workbook to a file
            string outputPath = "AutoFitRowsBlockDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
