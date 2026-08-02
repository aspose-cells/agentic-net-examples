// Title: Split an Excel worksheet into four quadrants with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, fills sample data, calls Worksheet.Split(), sets the first visible row of the bottom pane and the first visible column of the right pane via PaneCollection, and saves the file as WorksheetQuadrants.xlsx, producing four independent quadrants.
// Keywords: Aspose.Cells | C# | .NET | split worksheet panes | horizontal split | vertical split | first visible row bottom pane | first visible column right pane | Excel quadrants | PaneCollection
// Common Searches: Aspose.Cells split worksheet into four panes C# | how to set pane split position in Aspose.Cells | horizontal and vertical split Excel using Aspose.Cells | create quadrants in Excel workbook .NET | PaneCollection first visible row column example
// Developer Intent: The developer needs to divide a worksheet into four scrollable quadrants by applying both a horizontal and a vertical pane split.
// Use Cases: View separate sections of a large spreadsheet side‑by‑side without scrolling each area individually. | Freeze header rows and columns while allowing each quadrant to scroll independently for data analysis. | Prepare a layout where distinct data regions occupy fixed quadrants for printing or reporting.
// AI Prompts: Generate C# code using Aspose.Cells to split a worksheet at row 15 and column 8 and export the result to PDF. | Explain how to retrieve and modify PaneCollection properties after calling Worksheet.Split in Aspose.Cells. | Provide a step‑by‑step tutorial for creating four independent scrollable panes in an Excel file with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, fills sample data, calls Worksheet.Split(), sets the first visible row of the bottom pane and the first visible column of the right pane via PaneCollection, and saves the file as WorksheetQuadrants.xlsx, producing four independent quadrants.
    public class SplitWorksheetIntoQuadrants
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some data to visualize the split (optional)
                for (int row = 0; row < 30; row++)
                {
                    for (int col = 0; col < 15; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Split the window (horizontal and vertical split)
                worksheet.Split();

                // Obtain the pane collection to set split positions
                PaneCollection panes = worksheet.GetPanes();

                // Set the first visible row of the bottom pane (horizontal split)
                // This creates a split after row 10 (zero‑based index 9)
                panes.FirstVisibleRowOfBottomPane = 10;

                // Set the first visible column of the right pane (vertical split)
                // This creates a split after column 5 (zero‑based index 4)
                panes.FirstVisibleColumnOfRightPane = 5;

                // Save the workbook (lifecycle: save)
                workbook.Save("WorksheetQuadrants.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
