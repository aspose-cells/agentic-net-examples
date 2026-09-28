// Title: Delete all slicers from an Excel workbook using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, loops through every worksheet, clears each slicer collection, and saves the modified workbook. | Show how to use the Worksheet.Slicers.Clear method in Aspose.Cells to programmatically remove slicers linked to pivot tables across a workbook. | Provide a step‑by‑step Aspose.Cells example that deletes all slicers from a workbook and writes the result to a new file.
// Common Searches: asp.net aspose.cells remove slicers from all worksheets in an Excel file | c# clear slicer collection for each sheet using Aspose.Cells | how to delete pivot table slicers programmatically with Aspose.Cells .NET | iterate through workbook worksheets to remove slicers Aspose.Cells | save workbook after removing slicers Aspose.Cells C#
// Tags: Aspose.Cells slicer collection clear | delete all slicers workbook Aspose.Cells | worksheet level slicer removal Aspose.Cells | pivot table slicer cleanup Aspose.Cells | C# Aspose.Cells remove slicers

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// // Loads an .xlsx workbook, iterates each worksheet, calls sheet.Slicers.Clear() to delete all slicers, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your source file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear all slicers associated with the current worksheet
            // This removes slicers linked to any PivotTable on the sheet
            sheet.Slicers.Clear();
        }

        // Save the modified workbook to a new file (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
