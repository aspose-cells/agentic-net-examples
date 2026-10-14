// Title: How to delete all shapes from empty worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that removes every shape from worksheets that contain no data. | Create a method that scans a workbook, identifies empty sheets via MaxDataRow/MaxDataColumn, clears all shapes, and saves the result. | Write a routine that iterates through each worksheet, checks the first cell and used range, and deletes all shapes on empty sheets.
// Common Searches: Aspose.Cells C# delete shapes from blank worksheets | how to clear drawings on empty Excel sheets using .NET | remove all shapes from worksheets with no data programmatically
// Tags: remove shapes Aspose.Cells worksheet | detect empty worksheet MaxDataRow C# | remove drawings from blank Excel sheets Aspose.Cells | strip shapes from unused worksheets .NET | optimize workbook by removing shapes Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook, checks each worksheet for emptiness using MaxDataRow, MaxDataColumn, and the first cell value, removes all shapes from those empty sheets by iterating backwards, and saves the cleaned workbook.
class RemoveShapesFromEmptySheets
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine if the worksheet is empty based on cell data
            // An empty sheet has no used rows/columns and the first cell is null
            bool isEmpty = sheet.Cells.MaxDataRow == 0 &&
                           sheet.Cells.MaxDataColumn == 0 &&
                           sheet.Cells[0, 0].Value == null;

            if (isEmpty)
            {
                // Remove all shapes from the empty worksheet
                // Iterate backwards to avoid index shifting issues
                for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
                {
                    sheet.Shapes.RemoveAt(i);
                }
            }
        }

        // Save the modified workbook (replace with desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
