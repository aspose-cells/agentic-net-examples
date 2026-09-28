// Title: How to count the number of worksheets that contain shapes in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# function that loads an .xlsx file with Aspose.Cells and returns the count of worksheets that have at least one shape. | Show how to iterate through a Workbook's Worksheets collection and increment a counter when the Shapes collection size is greater than zero. | Provide sample code that prints the names of all sheets containing shapes and then displays the total number of such sheets.
// Common Searches: Aspose.Cells .NET how to find worksheets that include drawings or charts | C# get number of Excel sheets with any shape objects using Aspose.Cells API | sample code to list Excel worksheet names that contain shapes with Aspose.Cells | determine which worksheets in a workbook have shape collections populated in C#
// Tags: worksheet shape detection Aspose.Cells | count sheets with drawings .NET | enumerate shapes per worksheet using Aspose.Cells | Excel workbook shape analysis C# | report worksheets containing graphics Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel file with Aspose.Cells, loops through each worksheet, checks the Shapes collection count, increments a counter for worksheets that have shapes, and writes the total count (and optionally the sheet names) to the console.
class Program
{
    static void Main(string[] args)
    {
        // Path to the workbook file (modify as needed)
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        int worksheetsWithShapes = 0;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet contains any shapes
            if (sheet.Shapes.Count > 0)
            {
                worksheetsWithShapes++;
            }
        }

        // Print the count of worksheets that contain shapes
        Console.WriteLine($"Worksheets containing shapes: {worksheetsWithShapes}");
    }
}
