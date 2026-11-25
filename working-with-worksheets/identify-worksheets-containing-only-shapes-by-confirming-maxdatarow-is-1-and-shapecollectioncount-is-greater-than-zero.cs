// Title: Find Excel worksheets that contain only shapes (no cell data) using Aspose.Cells for .NET
// Common Searches: how to list worksheets that have only drawings and no cell values with Aspose.Cells C# | C# Aspose.Cells detect worksheets with shapes but no data | using Cells.MaxDataRow and Worksheet.Shapes to identify shape‑only sheets in .NET | Aspose.Cells find Excel sheets that contain only pictures and charts without any cell content | programmatically check for worksheets that contain only shapes in an Excel workbook using Aspose.Cells
// Tags: shape‑only worksheet detection Aspose.Cells | MaxDataRow -1 empty cells check | Worksheet.Shapes count evaluation | filter worksheets without cell data C# | identify Excel sheets with only drawings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace IdentifyShapeOnlyWorksheets
{
    // The example loads a workbook, iterates each worksheet, and reports those where Cells.MaxDataRow equals -1 (no cell data) while Worksheet.Shapes.Count is greater than zero, indicating the sheet contains only shapes. It prints the worksheet name and shape count or a message if none are found, with error handling for missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: IdentifyShapeOnlyWorksheets <path-to-workbook>");
                return;
            }

            string workbookPath = args[0];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found – \"{workbookPath}\"");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Flag to indicate if any matching worksheet was found
                bool anyMatch = false;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // MaxDataRow returns -1 when there is no cell containing data
                    int maxDataRow = sheet.Cells.MaxDataRow;

                    // ShapeCollection.Count gives the number of shapes on the sheet
                    int shapeCount = sheet.Shapes.Count;

                    // Identify worksheets that have only shapes (no data)
                    if (maxDataRow == -1 && shapeCount > 0)
                    {
                        anyMatch = true;
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only shapes (ShapeCount={shapeCount}).");
                    }
                }

                if (!anyMatch)
                {
                    Console.WriteLine("No worksheets containing only shapes were found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred while processing the workbook: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }
    }
}
