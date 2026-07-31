// Title: Aspose.Cells C# – Remove all shapes from empty worksheets
// Description: Loads a workbook, checks each worksheet for data using MaxDataRow/MaxDataColumn, removes every drawing object (shapes, pictures, charts, etc.) from sheets that contain no data, and saves the cleaned file. Ideal for shrinking file size and preparing templates.
// Keywords: Aspose.Cells remove shapes C# | delete drawing objects empty worksheet | RemoveAllDrawingObjects example | check worksheet emptiness Aspose.Cells | clean Excel workbook Aspose.Cells .NET | optimize workbook size Aspose.Cells | strip graphics from blank sheets
// Common Searches: How to delete all shapes from empty worksheets using Aspose.Cells for .NET | Remove drawing objects from blank Excel tabs C# | Identify empty worksheet Aspose.Cells and clear graphics | Aspose.Cells RemoveAllDrawingObjects on empty sheets | Reduce Excel file size by removing pictures from unused sheets
// Developer Intent: Programmatically clear every drawing object from worksheets that have no cell data.
// Use Cases: Prepare a distribution‑ready template by stripping graphics from unused tabs. | Archive reports while minimizing file size by eliminating drawings on blank sheets. | Pre‑process user‑uploaded Excel files in a web service to remove unnecessary graphics from empty worksheets.
// AI Prompts: Write C# code with Aspose.Cells that removes all drawing objects from worksheets lacking data rows or columns. | Suggest an alternative way to detect empty worksheets and delete only picture objects, preserving charts, in Aspose.Cells. | Show how to log the names of worksheets from which drawing objects were removed using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, checks each worksheet for data using MaxDataRow/MaxDataColumn, removes every drawing object (shapes, pictures, charts, etc.) from sheets that contain no data, and saves the cleaned file. Ideal for shrinking file size and preparing templates.
    class RemoveShapesFromEmptySheets
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet contains any data.
                // MaxDataRow and MaxDataColumn return -1 when there is no data.
                bool isEmpty = sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0;

                if (isEmpty)
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
