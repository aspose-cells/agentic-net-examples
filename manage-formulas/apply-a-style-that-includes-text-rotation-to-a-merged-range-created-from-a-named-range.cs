// Title: Apply a 45‑degree text rotation to a merged range derived from a named range with Aspose.Cells for .NET
// AI Prompts: Create a named range, merge its cells, and apply a 45° rotated style with centered alignment using Aspose.Cells in C#. | Define a StyleFlag that applies all attributes and use it to set text rotation on a merged range created from a named range. | Generate an Excel workbook where the merged area "MyRange" displays text rotated 45 degrees and save it as MergedRangeWithRotation.xlsx.
// Common Searches: Aspose.Cells C# rotate text in a merged range created from a named range | how to apply a style with rotation to a merged cell block using Aspose.Cells | set text orientation for merged cells after merging a named range in .NET | example of using StyleFlag to apply rotation to a merged range in Aspose.Cells
// Tags: merged range text rotation Aspose.Cells | style flag application to named range Aspose.Cells | named range merging with rotated style Aspose.Cells | create rotated cell style Aspose.Cells C# | export workbook with rotated merged cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // The example creates a workbook, defines a named range A1:C3, merges the cells, builds a style with a 45‑degree rotation and centered alignment, applies the style to the merged range via a StyleFlag, and saves the file as MergedRangeWithRotation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 1. Define a named range (A1:C3) and give it a name "MyRange"
                int startRow = 0;          // Row 1 (zero‑based)
                int startColumn = 0;       // Column A
                int rowCount = 3;          // 3 rows
                int columnCount = 3;       // 3 columns

                AsposeRange namedRange = sheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);
                namedRange.Name = "MyRange";

                // -----------------------------------------------------------------
                // 2. Merge the cells that belong to the named range
                sheet.Cells.Merge(namedRange.FirstRow, namedRange.FirstColumn, namedRange.RowCount, namedRange.ColumnCount);

                // -----------------------------------------------------------------
                // 3. Create a style that includes text rotation
                Style rotatedStyle = workbook.CreateStyle();
                rotatedStyle.RotationAngle = 45;                     // Rotate text 45 degrees
                rotatedStyle.HorizontalAlignment = TextAlignmentType.Center;
                rotatedStyle.VerticalAlignment = TextAlignmentType.Center;

                // Define which style attributes should be applied (apply all to include rotation)
                StyleFlag styleFlag = new StyleFlag
                {
                    All = true
                };

                // Apply the style to the merged range
                namedRange.ApplyStyle(rotatedStyle, styleFlag);

                // -----------------------------------------------------------------
                // 4. Save the workbook
                string outputPath = "MergedRangeWithRotation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
