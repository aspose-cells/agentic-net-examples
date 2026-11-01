// Title: Count worksheets that contain shapes in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that returns the total number of worksheets having at least one shape in a specified .xlsx file. | Write a method that accepts a workbook path and prints both the count and the names of sheets that contain any drawings with Aspose.Cells. | Create a console program that loads an Excel workbook and reports how many worksheets include shapes, leveraging the Shapes collection in Aspose.Cells.
// Common Searches: asp.net count worksheets with drawings using Aspose.Cells | c# get number of sheets that contain shapes in an Excel file Aspose.Cells | how to list worksheet names that have shapes with Aspose.Cells .NET | Aspose.Cells shape detection per worksheet example
// Tags: count worksheets with shapes Aspose.Cells | list sheets containing drawings .NET | worksheet shape enumeration Aspose.Cells | excel workbook shape detection C# | report sheets having shapes using Aspose.Cells

using System;
using Aspose.Cells;

// // Loads a workbook and prints the number of worksheets that contain at least one shape using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "input.xlsx";

        // Load the workbook
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

        // Output the count
        Console.WriteLine($"Worksheets containing shapes: {worksheetsWithShapes}");
    }
}
