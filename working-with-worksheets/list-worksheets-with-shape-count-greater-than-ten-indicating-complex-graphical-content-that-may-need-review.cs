// Title: Identify Excel worksheets with more than ten shapes using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that prints each worksheet name and its shape count when the count exceeds ten. | Create a method that returns a collection of worksheet names where the number of shapes is greater than ten using Aspose.Cells. | Adapt the sample to log the worksheet name together with the total shapes for all sheets that have more than ten drawings.
// Common Searches: aspnet how to list worksheets that have over 10 shapes with Aspose.Cells | C# Aspose.Cells count drawings per worksheet and filter results | retrieve Excel sheet names where shape count is greater than ten using Aspose.Cells | filter worksheets by number of shapes in a workbook Aspose.Cells .NET | find complex graphical content in Excel sheets with Aspose.Cells C#
// Tags: Aspose.Cells count worksheet shapes | C# filter worksheets by shape count | Aspose.Cells retrieve sheets with many drawings | Excel workbook shape analysis Aspose.Cells | list worksheets exceeding shape threshold .NET

using System;
using Aspose.Cells;

// // Loads an Excel workbook with Aspose.Cells, iterates through each worksheet, checks the Shapes collection count, and writes to the console the worksheet name and shape count for any sheet that contains more than ten shapes.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the number of shapes on the current worksheet
            int shapeCount = sheet.Shapes.Count;

            // If the worksheet has more than ten shapes, output its name and count
            if (shapeCount > 10)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {shapeCount} shapes.");
            }
        }
    }
}
