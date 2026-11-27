// Title: C# code to identify and flag Excel worksheets that contain only shapes (no cell data) using Aspose.Cells
// Common Searches: Aspose.Cells C# find worksheets that have shapes but no cell values | how to detect shape‑only sheets in an Excel file with Aspose.Cells for .NET | C# program to list worksheets where MaxDataRow and MaxDataColumn are -1 and Shapes count > 0 | flag Excel worksheets that contain only drawings using Aspose.Cells API | scan workbook for worksheets without data cells but with drawing objects in C#
// Tags: detect shape‑only worksheets Aspose.Cells | flag worksheets without cell data C# | shape collection count check Aspose.Cells | maxdatarow maxdatacolumn -1 detection | scan Excel workbook for drawing‑only sheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeOnlyWorksheetFlagger
{
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and flags any sheet where the Shapes collection count is greater than zero while Cells.MaxDataRow and Cells.MaxDataColumn both return -1, indicating the worksheet contains only shapes and no cell data.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args == null || args.Length == 0)
                {
                    Console.WriteLine("Usage: ShapeOnlyWorksheetFlagger <path-to-workbook>");
                    return;
                }

                string workbookPath = args[0];

                if (string.IsNullOrWhiteSpace(workbookPath))
                {
                    Console.WriteLine("Error: The workbook path is empty.");
                    return;
                }

                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: File not found – \"{workbookPath}\"");
                    return;
                }

                // Load the workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(workbookPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Error loading workbook \"{workbookPath}\": {loadEx.Message}");
                    return;
                }

                // Iterate through worksheets
                bool anyFlagged = false;
                Console.WriteLine($"Scanning workbook \"{Path.GetFileName(workbookPath)}\" for shape‑only worksheets...");

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Retrieve shape collection and cell data extents
                    ShapeCollection shapes = sheet.Shapes;
                    int shapeCount = shapes.Count;

                    // Cells.MaxDataRow / MaxDataColumn return -1 when no data exists
                    int maxDataRow = sheet.Cells.MaxDataRow;
                    int maxDataColumn = sheet.Cells.MaxDataColumn;

                    // Determine if the sheet is shape‑only
                    bool isShapeOnly = shapeCount > 0 && maxDataRow == -1 && maxDataColumn == -1;

                    if (isShapeOnly)
                    {
                        anyFlagged = true;
                        Console.WriteLine($"[FLAGGED] Worksheet \"{sheet.Name}\" (Index {sheet.Index}) contains {shapeCount} shape(s) but no data cells.");
                    }
                    else
                    {
                        // Optional verbose output – uncomment if needed
                        //Console.WriteLine($"Worksheet \"{sheet.Name}\": Shapes={shapeCount}, MaxDataRow={maxDataRow}, MaxDataColumn={maxDataColumn}");
                    }
                }

                if (!anyFlagged)
                {
                    Console.WriteLine("No shape‑only worksheets were found.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exception and report it
                Console.WriteLine($"Unexpected error: {ex.GetType().Name} – {ex.Message}");
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
