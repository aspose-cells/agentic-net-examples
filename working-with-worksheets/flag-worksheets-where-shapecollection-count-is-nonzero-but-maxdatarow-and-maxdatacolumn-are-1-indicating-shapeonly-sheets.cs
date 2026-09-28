// Title: Identify and flag shape‑only worksheets in an Excel file using Aspose.Cells for .NET by adding a comment and coloring the tab
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, detects sheets where Shapes.Count > 0 and Cells.MaxDataRow/MaxDataColumn are -1, then adds a comment "Shape‑only sheet" to cell A1 and sets the worksheet tab color to yellow. | Create a reusable C# method using Aspose.Cells that marks shape‑only worksheets by inserting a comment in A1, applying a yellow tab highlight, and returns the updated workbook.
// Common Searches: aspnet how to find Excel worksheets that contain only drawings with Aspose.Cells | c# mark shape‑only sheets in a workbook by adding a comment and changing tab color | using Aspose.Cells detect worksheets with Shapes collection but no data rows | flag Excel sheets that have drawings but no cell data in .NET
// Tags: detect shape‑only worksheets Aspose.Cells | add comment to worksheet cell A1 C# | set worksheet tab color Aspose.Cells | check MaxDataRow MaxDataColumn -1 Aspose.Cells | shape collection count > 0 Excel .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet, and when a sheet has Shapes.Count > 0 while Cells.MaxDataRow and Cells.MaxDataColumn are -1 (indicating no data cells), it adds a comment "Shape‑only sheet" to cell A1, colors the worksheet tab yellow, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Condition: sheet contains shapes but has no data rows/columns
                bool hasShapes = sheet.Shapes.Count > 0;
                bool hasNoData = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

                if (hasShapes && hasNoData)
                {
                    // Flag the sheet by adding a comment to cell A1
                    int commentIndex = sheet.Comments.Add("A1");
                    Comment comment = sheet.Comments[commentIndex];
                    comment.Note = "Shape‑only sheet";

                    // Optionally set the worksheet tab color to highlight it
                    sheet.TabColor = Color.Yellow;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
