// Title: Aspose.Cells C# – List worksheets containing more than 10 drawing shapes
// Description: Load an Excel workbook with Aspose.Cells for .NET, iterate each worksheet, check the Shapes collection, and output the names of sheets that have over ten shapes—an indicator of complex graphical content that may affect performance or require review.
// Keywords: Aspose.Cells shape count | C# worksheet shapes | list worksheets with many shapes | filter Excel sheets by drawing objects | detect complex graphics in workbook | performance impact shapes Aspose | Excel shape threshold | Aspose.Cells .NET graphics analysis
// Common Searches: C# Aspose.Cells list worksheets with more than 10 shapes | How to count shapes per worksheet using Aspose.Cells | Identify Excel sheets with high number of drawing objects | Filter worksheets by shape count Aspose.Cells | Find complex graphical content in Excel workbook C#
// Developer Intent: Retrieve the names and shape counts of worksheets that exceed a ten‑shape threshold.
// Use Cases: Audit workbooks for sheets with heavy graphics before publishing. | Spot performance‑critical worksheets that may slow loading or rendering. | Generate a report of sheets needing graphic cleanup or shape optimization.
// AI Prompts: Create a reusable method that returns a list of worksheet names where sheet.Shapes.Count > 10 using Aspose.Cells. | Adapt the sample to write the worksheet names and shape counts to a CSV file instead of the console. | Build a utility class that logs high‑shape worksheets and optionally removes shapes beyond a configurable limit.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load an Excel workbook with Aspose.Cells for .NET, iterate each worksheet, check the Shapes collection, and output the names of sheets that have over ten shapes—an indicator of complex graphical content that may affect performance or require review.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the number of drawing shapes in the current worksheet
            int shapeCount = sheet.Shapes.Count;

            // If the worksheet contains more than ten shapes, output its name and count
            if (shapeCount > 10)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" has {shapeCount} shapes (complex graphical content).");
            }
        }
    }
}
