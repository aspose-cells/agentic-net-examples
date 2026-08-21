// Title: C# – Set Fixed Column Width and Auto‑Fit Adjacent Column with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, write short text to column A and long text to column B, apply a fixed width to column A using Cells.SetColumnWidth, display the original width of column B, auto‑fit column B with Worksheet.AutoFitColumn, retrieve the adjusted width, and save the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | SetColumnWidth | AutoFitColumn | GetColumnWidth | fixed column width | auto fit column | Excel column sizing | worksheet column width | Aspose.Cells example
// Common Searches: Aspose.Cells set column width C# | How to auto‑fit a column after fixing another column in Aspose.Cells | Get column width before and after AutoFit Aspose.Cells | C# example for fixed vs auto‑fit column widths in Excel | Worksheet.AutoFitColumn usage Aspose.Cells .NET
// Developer Intent: Apply a specific width to one column and automatically adjust the width of a neighboring column for layout comparison.
// Use Cases: Design a report where header columns stay at a constant width while data columns expand to fit content. | Create Excel sheets that mix fixed‑width and auto‑fit columns to maintain readability across varied text lengths. | Log column dimensions before and after AutoFit to troubleshoot spreadsheet layout issues.
// AI Prompts: Generate C# code that sets column A to 25 characters wide and then auto‑fits column B using Aspose.Cells. | Show how to retrieve and print column widths before and after calling Worksheet.AutoFitColumn in a .NET workbook. | Explain the algorithm Aspose.Cells uses to calculate optimal column width during AutoFitColumn.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, write short text to column A and long text to column B, apply a fixed width to column A using Cells.SetColumnWidth, display the original width of column B, auto‑fit column B with Worksheet.AutoFitColumn, retrieve the adjusted width, and save the file as an Excel workbook.
    public class SetAndAutoFitColumnDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate column A with short text and column B with long text
                cells["A1"].PutValue("Short");
                cells["A2"].PutValue("Data");
                cells["B1"].PutValue("This is a much longer piece of text that will require column width adjustment");
                cells["B2"].PutValue("Another long text entry to demonstrate AutoFitColumn functionality");

                // Set a fixed width for column A (index 0) – 20 characters
                cells.SetColumnWidth(0, 20.0);
                Console.WriteLine($"Column A width set to: {cells.GetColumnWidth(0)} characters");

                // Display column B width before auto‑fit
                Console.WriteLine($"Column B width before AutoFit: {cells.GetColumnWidth(1)} characters");

                // Auto‑fit column B (index 1) based on its content
                worksheet.AutoFitColumn(1);
                Console.WriteLine($"Column B width after AutoFit: {cells.GetColumnWidth(1)} characters");

                // Save the workbook
                string outputPath = "SetAndAutoFitColumnDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            SetAndAutoFitColumnDemo.Run();
        }
    }
}
