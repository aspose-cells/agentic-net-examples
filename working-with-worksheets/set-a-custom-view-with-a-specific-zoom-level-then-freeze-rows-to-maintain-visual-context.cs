// Title: Set a custom zoom level and freeze top rows/columns in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet zoom to 150% and freeze first two rows | C# example for FreezePanes after applying a custom view with Aspose.Cells | How to programmatically freeze panes and adjust zoom in an Excel file using Aspose.Cells .NET | Apply custom view and freeze top rows in an Aspose.Cells workbook with C# | Aspose.Cells .NET freeze top rows and columns while setting view type
// Tags: worksheet zoom Aspose.Cells .NET | freeze panes Aspose.Cells C# | custom view Excel Aspose.Cells | set view type Aspose.Cells workbook | freeze top rows columns Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomViewAndFreeze
{
    // The sample loads an existing Excel file (or creates a new workbook), renames the first worksheet, sets its Zoom property to 150 %, optionally sets the view type, freezes the first two rows and the first column using FreezePanes with cell "B3", and saves the result as an XLSX file.
    class Program
    {
        /// <param name="args">
        /// args[0] (optional) – path to an existing Excel file to modify.
        /// args[1] (optional) – path where the resulting file will be saved.
        /// </param>
        static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "output_custom_view.xlsx";

                // Ensure the input file exists; if not, create a simple workbook.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                    // Populate a few rows so the freeze effect is visible.
                    Worksheet wsInit = workbook.Worksheets[0];
                    for (int row = 0; row < 30; row++)
                    {
                        for (int col = 0; col < 10; col++)
                        {
                            wsInit.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                        }
                    }
                }

                // Work with the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "CustomViewSheet";

                // -------------------------------------------------
                // 1. Set a custom view – here we set the zoom level.
                // -------------------------------------------------
                // Valid zoom values are between 10 and 400 (percentage).
                int desiredZoom = 150; // 150%
                worksheet.Zoom = desiredZoom;
                Console.WriteLine($"Worksheet zoom set to {desiredZoom}%.");

                // Optionally, you can change the view type (e.g., Normal, PageLayout, PageBreakPreview)
                // worksheet.ViewType = ViewType.Normal; // default is Normal

                // -------------------------------------------------
                // 2. Freeze rows/columns to keep headers visible.
                // -------------------------------------------------
                // Example: Freeze the top 2 rows and the first column.
                // FreezePanes(string cellName, int freezedRows, int freezedColumns)
                // The cellName indicates the cell just below/right of the frozen pane.
                // For freezing rows 1‑2 and column A, use cell "B3".
                string freezeCell = "B3";
                int frozenRows = 2;      // rows 1 and 2 will stay visible
                int frozenColumns = 1;   // column A will stay visible

                worksheet.FreezePanes(freezeCell, frozenRows, frozenColumns);
                Console.WriteLine($"Frozen panes set at cell \"{freezeCell}\" (Rows: {frozenRows}, Columns: {frozenColumns}).");

                // -------------------------------------------------
                // Save the modified workbook
                // -------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while processing the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
