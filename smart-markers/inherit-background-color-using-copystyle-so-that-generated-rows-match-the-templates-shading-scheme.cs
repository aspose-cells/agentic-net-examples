// Title: Copy background shading from a template row to inserted rows using Aspose.Cells CopyStyle in C#
// Common Searches: Aspose.Cells copy row background color to new rows C# | How to preserve template row shading when inserting rows with Aspose.Cells | C# insert rows and inherit cell style using SetStyle Aspose.Cells | CopyStyle method to duplicate row formatting in an Excel workbook C# | Apply template row formatting to multiple inserted rows Aspose.Cells
// Tags: copy row style Aspose.Cells C# | inherit background shading inserted rows | template row formatting Aspose.Cells | setstyle bulk cell styling | copystyle method Aspose.Cells C# | insert rows preserve Excel formatting Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, inserts a set of blank rows at a chosen position, and uses GetStyle/SetStyle (CopyStyle) to transfer the full style—including background color—from a designated template row to every cell of the newly added rows, then saves the result with robust file‑existence checks and per‑cell error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook
            var workbook = new Workbook(templatePath);
            var sheet = workbook.Worksheets[0];

            // Row that contains the desired background shading (0‑based)
            int templateRowIndex = 1; // adjust as needed
            // Position where new rows will be inserted (0‑based)
            int insertAtRow = 5;
            // Number of rows to add
            int rowsToAdd = 4;

            // Insert blank rows
            sheet.Cells.InsertRows(insertAtRow, rowsToAdd);

            // Determine the number of columns to copy styles for
            int maxColumn = sheet.Cells.MaxColumn + 1;

            // Copy style from the template row to each newly inserted row
            for (int i = 0; i < rowsToAdd; i++)
            {
                int targetRowIndex = insertAtRow + i;
                for (int col = 0; col < maxColumn; col++)
                {
                    try
                    {
                        var templateCell = sheet.Cells[templateRowIndex, col];
                        var targetCell = sheet.Cells[targetRowIndex, col];
                        // Copy the entire style (including background color)
                        targetCell.SetStyle(templateCell.GetStyle());
                    }
                    catch (Exception ex)
                    {
                        // Log but continue with other cells
                        Console.WriteLine($"Failed to copy style for cell [{targetRowIndex}, {col}]: {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
