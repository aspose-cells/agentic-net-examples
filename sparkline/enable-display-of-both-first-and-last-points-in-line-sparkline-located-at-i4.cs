// Title: Enable first and last point markers on a line sparkline in cell I4 using Aspose.Cells for .NET
// AI Prompts: Set the ShowFirstPoint and ShowLastPoint properties of the sparkline group that contains the line sparkline in cell I4, then save the workbook. | Load an existing Excel file, locate the sparkline group for cell I4, turn on first and last point markers for the line sparkline, and write the changes to a new file with C# and Aspose.Cells.
// Common Searches: Aspose.Cells C# how to show first point on a line sparkline in a specific cell | enable both first and last markers for a sparkline at I4 using Aspose.Cells .NET | modify sparkline group properties to display end points in Excel with Aspose.Cells
// Tags: Aspose.Cells sparkline first point marker | Aspose.Cells sparkline last point marker | C# edit sparkline group properties Excel | line sparkline end point markers .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, finds the sparkline group, enables ShowFirstPoint and ShowLastPoint to display the first and last markers on the line sparkline in cell I4, and saves the modified workbook.
class SparklineExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Check for existing sparkline groups
            if (sheet.SparklineGroups.Count > 0)
            {
                // Get the first sparkline group
                var sparklineGroup = sheet.SparklineGroups[0];

                // Enable display of first and last points
                sparklineGroup.ShowFirstPoint = true;
                sparklineGroup.ShowLastPoint = true;
            }
            else
            {
                Console.WriteLine("No sparkline groups found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
