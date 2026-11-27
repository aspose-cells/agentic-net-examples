// Title: How to count worksheets that contain shapes in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells count worksheets that have shapes | Get number of Excel sheets with drawings using Aspose.Cells | Determine which worksheets contain shapes in a workbook with Aspose.Cells | Iterate through worksheets to find shape collections in Aspose.Cells C# | Report total worksheets containing drawings in an Excel file using Aspose.Cells
// Tags: worksheet shape detection Aspose.Cells | enumerate shape collections in Excel worksheets | count worksheets with drawings .NET | Aspose.Cells shape count per worksheet | excel workbook shape analysis C# | iterate worksheets for shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WorksheetShapeCounter
{
    // The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, checks the Shapes collection, increments a counter when a worksheet contains at least one shape, and prints the total number of worksheets that have shapes.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine workbook path
                string workbookPath;
                if (args.Length > 0)
                {
                    workbookPath = args[0];
                }
                else
                {
                    workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.xlsx");
                }

                // Validate file existence
                if (!File.Exists(workbookPath))
                {
                    Console.Error.WriteLine($"Error: Workbook file not found at \"{workbookPath}\".");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Count worksheets that contain shapes
                int worksheetsWithShapes = 0;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // ShapeCollection is accessed via Worksheet.Shapes
                    ShapeCollection shapes = sheet.Shapes;
                    if (shapes != null && shapes.Count > 0)
                    {
                        worksheetsWithShapes++;
                    }
                }

                // Output the result
                Console.WriteLine($"Total worksheets containing shapes: {worksheetsWithShapes}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
