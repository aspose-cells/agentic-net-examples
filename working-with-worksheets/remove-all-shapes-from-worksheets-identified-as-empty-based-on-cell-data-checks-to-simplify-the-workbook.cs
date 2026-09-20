// Title: Remove all shapes from empty worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, finds worksheets without any cell data, and clears their Shapes collection. | Show how to use MaxDataRow and MaxDataColumn to determine empty sheets and then remove all drawings from those sheets in Aspose.Cells. | Provide a reusable method that iterates through a workbook, detects blank worksheets, and calls Shapes.Clear() for each.
// Common Searches: Aspose.Cells C# remove drawings from worksheets that have no data | how to clear shapes collection on blank Excel sheets using .NET | detect empty worksheet with MaxDataRow MaxDataColumn Aspose.Cells | C# code to delete all shapes from empty worksheets in an Excel workbook | automate removal of graphics from blank sheets with Aspose.Cells
// Tags: clear shapes collection Aspose.Cells | empty worksheet detection MaxDataRow Aspose.Cells | remove drawings from blank Excel sheet C# | iterate worksheets and clear shapes Aspose.Cells | use MaxDataColumn to identify empty sheet .NET

using System;
using Aspose.Cells;

// Loads a workbook, checks each worksheet for data using MaxDataRow and MaxDataColumn, clears all shapes from sheets with no data, and saves the updated workbook.
class RemoveShapesFromEmptySheets
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine if the worksheet is empty (no data in any cell)
            // MaxDataRow and MaxDataColumn return -1 when there is no data
            bool isEmpty = sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0;

            if (isEmpty)
            {
                // Remove all shapes from the empty worksheet
                // Clear the Shapes collection
                sheet.Shapes.Clear();
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
