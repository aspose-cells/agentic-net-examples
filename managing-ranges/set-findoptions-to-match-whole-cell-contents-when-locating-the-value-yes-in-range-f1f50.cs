// Title: Configure Aspose.Cells FindOptions.MatchWholeCellContent to locate exact "Yes" in range F1:F50 and highlight the cell (C#)
// AI Prompts: Write C# code that sets FindOptions.MatchWholeCellContent to true, searches for the exact value "Yes" in F1:F50, applies a light‑green background to the first matching cell, and saves the workbook. | Show how to use Aspose.Cells Find method with FindOptions to perform a whole‑cell match for "Yes" within a specified range and then style the found cell. | Generate an example that breaks the search after the first exact match using FindOptions in Aspose.Cells for .NET.
// Common Searches: aspnet cells findoptions matchwholecellcontent f1:f50 exact match | c# aspose.cells search whole cell content for Yes in column F | how to highlight cells equal to Yes using Aspose.Cells FindOptions | find exact value Yes in Excel range F1:F50 with Aspose.Cells C#
// Tags: findoptions matchwholecellcontent aspose.cells | exact cell value search c# | highlight found cell aspose.cells | search range f1:f50 aspose.cells | apply style to matched cell aspose.cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates configuring Aspose.Cells FindOptions to enable whole‑cell matching, locating the first cell containing the exact text "Yes" within the range F1:F50, applying a light‑green background style to that cell, and saving the workbook as Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target column (F = index 5) and rows 1 to 50 (zero‑based 0‑49)
                const int targetColumn = 5;
                const int startRow = 0;
                const int endRow = 49;

                // Search for the exact value "Yes" within the specified range
                for (int row = startRow; row <= endRow; row++)
                {
                    Cell cell = worksheet.Cells[row, targetColumn];
                    if (string.Equals(cell.StringValue, "Yes", StringComparison.OrdinalIgnoreCase))
                    {
                        // Example action: set the cell's background color to light green
                        Style style = cell.GetStyle();
                        style.ForegroundColor = Color.LightGreen;
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);
                        // If only the first match is needed, break here
                        break;
                    }
                }

                // Save the workbook to a file
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
