// Title: Copy chart shapes between Excel workbooks while preserving formatting – Aspose.Cells C# example
// Description: A C# utility that loads a source Excel file, creates a blank destination workbook, copies each worksheet’s cells and styles, then transfers only chart shapes (MsoDrawingType.Chart) with their original row/column positions, size and formatting. The charts are added to the matching sheets and the result is saved as a new file. Demonstrates Aspose.Cells for .NET ShapeCollection handling, error checking and folder creation.
// Keywords: Aspose.Cells | C# | copy chart shapes | Excel chart duplication | preserve chart formatting | ShapeCollection | MsoDrawingType.Chart | Excel workbook cloning | chart object transfer | GitHub example
// Common Searches: copy chart shapes Aspose.Cells C# | preserve Excel chart formatting programmatically | transfer charts between workbooks .NET | Aspose.Cells copy only charts without drawings | C# example copy Excel chart objects | GitHub Aspose.Cells chart copy utility
// Developer Intent: Copy every chart object from a source Excel workbook to a new workbook, keeping the original layout, size and formatting.
// Use Cases: Create report files by reusing charts from a master template | Migrate legacy Excel charts into a new workbook structure without losing design | Build a consolidated dashboard by extracting chart shapes from multiple source files | Automate generation of client‑specific workbooks that share a common set of charts
// AI Prompts: Generate C# code that copies chart shapes from one workbook to another using Aspose.Cells while retaining position and style. | Explain how to extend the utility to also copy chart data series, axis settings, and embedded images. | Provide robust error handling for missing source files, worksheets without charts, and permission issues during save. | Suggest ways to log the copy process and report which charts were transferred.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ChartCopyUtilityDemo
{
    // A C# utility that loads a source Excel file, creates a blank destination workbook, copies each worksheet’s cells and styles, then transfers only chart shapes (MsoDrawingType.Chart) with their original row/column positions, size and formatting. The charts are added to the matching sheets and the result is saved as a new file. Demonstrates Aspose.Cells for .NET ShapeCollection handling, error checking and folder creation.
    public static class ChartCopyUtility
    {
        /// <param name="sourceFilePath">Path to the source Excel file.</param>
        /// <param name="destFilePath">Path where the destination Excel file will be saved.</param>
        public static void CopyCharts(string sourceFilePath, string destFilePath)
        {
            try
            {
                // Verify source file exists
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourceFilePath);

                // Create an empty destination workbook and clear default sheet
                Workbook destWorkbook = new Workbook();
                destWorkbook.Worksheets.Clear();

                // Iterate through each worksheet in the source workbook
                foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
                {
                    // Copy worksheet contents (cells, formats, etc.) without drawings
                    int destSheetIndex = destWorkbook.Worksheets.AddCopy(sourceSheet.Index);
                    Worksheet destSheet = destWorkbook.Worksheets[destSheetIndex];
                    destSheet.Name = sourceSheet.Name;

                    // Copy chart shapes from source to destination
                    ShapeCollection sourceShapes = sourceSheet.Shapes;
                    ShapeCollection destShapes = destSheet.Shapes;

                    foreach (Shape sourceShape in sourceShapes)
                    {
                        // Identify chart shapes (MsoDrawingType.Chart)
                        if ((MsoDrawingType)sourceShape.Type == MsoDrawingType.Chart)
                        {
                            // Preserve original position
                            int topRow = sourceShape.UpperLeftRow;
                            int top = sourceShape.Top;
                            int leftColumn = sourceShape.UpperLeftColumn;
                            int left = sourceShape.Left;

                            // Add a copy of the chart shape to the destination sheet
                            destShapes.AddCopy(sourceShape, topRow, top, leftColumn, left);
                        }
                    }
                }

                // Ensure destination directory exists
                string destDir = Path.GetDirectoryName(destFilePath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the destination workbook
                destWorkbook.Save(destFilePath);
                Console.WriteLine($"Destination workbook saved to: {destFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during chart copy: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourcePath = "SourceWorkbook.xlsx";
            string destinationPath = "DestinationWorkbook.xlsx";

            ChartCopyUtility.CopyCharts(sourcePath, destinationPath);

            Console.WriteLine("All chart shapes have been processed.");
        }
    }
}
