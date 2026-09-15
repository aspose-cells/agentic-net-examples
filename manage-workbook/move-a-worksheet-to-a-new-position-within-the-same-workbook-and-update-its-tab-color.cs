// Title: Move a worksheet to the first tab and set its tab color to LightBlue using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that locates a worksheet by name (with a fallback to index), moves it to the first tab, changes its TabColor to LightBlue, and saves the workbook. | Generate a .NET example that loads an existing .xlsx file, reorders a specific sheet to index 0, applies a LightBlue tab color, and writes the updated file using Aspose.Cells. | Create a C# program that checks for a sheet named "Sheet2", moves it to the leftmost tab, sets the tab background to LightBlue, and handles missing files gracefully with Aspose.Cells.
// Common Searches: how to move a worksheet to the first tab with Aspose.Cells in C# | Aspose.Cells set Excel sheet tab color programmatically .NET | reorder Excel sheets and change tab color using Aspose.Cells library | fallback to worksheet index when name not found Aspose.Cells example | save modified workbook after changing sheet order Aspose.Cells C#
// Tags: move worksheet to specific tab Aspose.Cells | set worksheet tab color LightBlue .NET | load and save Excel workbook Aspose.Cells | fallback worksheet selection by name or index Aspose | reorder sheets in .xlsx using Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', searches for a worksheet named 'Sheet2' (or uses the second sheet as a fallback), moves that worksheet to the first tab position, changes its tab color to LightBlue, and saves the modified workbook as 'output.xlsx', with error handling for missing files.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Attempt to get the worksheet by name; fallback to index if not found
                string targetSheetName = "Sheet2";
                Worksheet sheet = workbook.Worksheets[targetSheetName];

                if (sheet == null && workbook.Worksheets.Count > 1)
                {
                    // Use the second worksheet (index 1) as a fallback
                    sheet = workbook.Worksheets[1];
                }

                if (sheet == null)
                {
                    Console.WriteLine("Target worksheet not found.");
                    return;
                }

                // Move the worksheet to the first tab position
                sheet.MoveTo(0);

                // Set a new tab color for the worksheet
                sheet.TabColor = Color.LightBlue;

                // Save the modified workbook
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
