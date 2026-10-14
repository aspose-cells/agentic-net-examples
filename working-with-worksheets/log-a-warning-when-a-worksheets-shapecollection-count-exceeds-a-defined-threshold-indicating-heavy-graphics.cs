// Title: Log a warning when a worksheet's shape count exceeds a threshold using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, iterates all worksheets, and prints a warning for any worksheet whose Shapes.Count is greater than a specified limit. | Modify the HeavyGraphicsChecker example to read the shape limit from a configuration file and write the warnings to a log file instead of the console. | Create a reusable C# method that takes a Workbook and a maximum shape count, returns the names of worksheets that exceed the limit, and logs a warning using Aspose.Cells.
// Common Searches: how to detect worksheets with many shapes using Aspose.Cells C# | Aspose.Cells warning for excessive graphics in Excel worksheets | C# check shape count threshold in Excel workbook with Aspose.Cells | log performance warnings for heavy graphics in Excel files .NET
// Tags: worksheet shape count threshold Aspose.Cells | log warning for excessive Excel graphics .NET | detect heavy graphics in workbook using Aspose.Cells | validate shapes per worksheet Aspose.Cells | performance impact of many shapes Excel .NET

using System;
using Aspose.Cells;

// C# program that loads an Excel workbook with Aspose.Cells, iterates each worksheet, checks the number of shapes, and writes a console warning when the shape count exceeds a defined threshold.
class HeavyGraphicsChecker
{
    static void Main()
    {
        // Define the path to the Excel file
        string filePath = "input.xlsx";

        // Load the workbook using Aspose.Cells
        Workbook workbook = new Workbook(filePath);

        // Define the threshold for heavy graphics (number of shapes)
        int shapeThreshold = 50;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the count of shapes in the current worksheet
            int shapeCount = sheet.Shapes.Count;

            // If the count exceeds the defined threshold, log a warning
            if (shapeCount > shapeThreshold)
            {
                Console.WriteLine($"Warning: Worksheet \"{sheet.Name}\" contains {shapeCount} shapes, which exceeds the threshold of {shapeThreshold}. Consider reducing graphics to improve performance.");
            }
        }

        // Optionally, save the workbook if any modifications were made
        // workbook.Save("output.xlsx");
    }
}
