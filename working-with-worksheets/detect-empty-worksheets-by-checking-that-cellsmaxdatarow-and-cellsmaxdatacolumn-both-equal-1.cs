// Title: Detect blank worksheets in an Excel workbook with Aspose.Cells for .NET by evaluating MaxDataRow and MaxDataColumn
// AI Prompts: Write C# using Aspose.Cells that iterates over every sheet and outputs the name when both Cells.MaxDataRow and Cells.MaxDataColumn return -1. | Implement a .NET function that returns a collection of sheet names that contain no data, based on MaxDataRow/MaxDataColumn checks. | Adjust the example to bypass any further actions for sheets identified as having no data and optionally delete them from the workbook.
// Common Searches: Aspose.Cells .NET determine if a sheet is empty using MaxDataRow | C# list worksheets without any cells populated in an .xlsx file | How to find blank Excel sheets with Aspose.Cells API | Detect worksheets lacking data in a workbook using Aspose.Cells for C#
// Tags: blank worksheet detection Aspose.Cells | MaxDataRow MaxDataColumn validation .NET | enumerate sheets with no data C# | skip empty sheets Aspose.Cells processing | remove worksheets without content Aspose API

using System;
using Aspose.Cells;

// Loads an Excel workbook, iterates through each worksheet, and reports whether the sheet is empty by checking that Cells.MaxDataRow and Cells.MaxDataColumn are both -1.
class DetectEmptyWorksheets
{
    static void Main(string[] args)
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if both MaxDataRow and MaxDataColumn are -1, indicating an empty sheet
            if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" is empty.");
            }
            else
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains data.");
            }
        }

        // Optionally, save the workbook if any modifications were made
        // workbook.Save("output.xlsx");
    }
}
