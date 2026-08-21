// Title: Remove All Shapes and Drawing Objects from Empty Worksheets with Aspose.Cells for .NET
// Description: Learn how to programmatically identify worksheets that contain no cell data (MaxDataRow and MaxDataColumn are -1) and delete every drawing object—shapes, pictures, charts, etc.—using Aspose.Cells' RemoveAllDrawingObjects method, then save the cleaned workbook.
// Keywords: Aspose.Cells remove shapes C# | delete drawing objects empty worksheet | RemoveAllDrawingObjects example | clear graphics from blank Excel sheets | Aspose.Cells worksheet cleanup .NET
// Common Searches: how to delete all shapes from empty worksheets Aspose.Cells | remove drawing objects from blank Excel tabs C# | check if worksheet is empty Aspose.Cells and clear drawings | Aspose.Cells RemoveAllDrawingObjects usage | clean empty sheets in Excel with Aspose.Cells
// Developer Intent: Detect worksheets without any cell content and purge all drawing objects from those sheets.
// Use Cases: Reduce file size before distributing a workbook by stripping graphics from unused tabs. | Automate preprocessing of user‑uploaded Excel files so blank sheets contain no leftover charts or images. | Enforce template compliance by ensuring placeholder worksheets are completely empty of drawings.
// AI Prompts: Generate a C# snippet using Aspose.Cells that loops through all worksheets, checks for emptiness, and calls RemoveAllDrawingObjects on empty sheets. | Explain the impact of RemoveAllDrawingObjects on different drawing types (shapes, pictures, charts) in an empty worksheet. | Provide step‑by‑step guidance for cleaning a workbook of all graphics on blank worksheets before saving.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Learn how to programmatically identify worksheets that contain no cell data (MaxDataRow and MaxDataColumn are -1) and delete every drawing object—shapes, pictures, charts, etc.—using Aspose.Cells' RemoveAllDrawingObjects method, then save the cleaned workbook.
    class RemoveShapesFromEmptyWorksheets
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet is empty (no data in any cell)
                // MaxDataRow returns -1 when there is no data
                if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1)
                {
                    // Remove all drawing objects (shapes, pictures, charts, etc.) from the empty worksheet
                    sheet.RemoveAllDrawingObjects();
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
