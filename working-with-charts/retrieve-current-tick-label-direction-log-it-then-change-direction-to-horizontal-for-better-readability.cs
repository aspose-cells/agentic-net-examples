// Title: Read current tick label rotation of an Excel chart and set it to horizontal using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that retrieves the CategoryAxis.TickLabelRotationAngle of the first chart, prints the angle, and updates it to 0 degrees for horizontal labels. | Demonstrate how to use reflection to modify the TickLabelRotationAngle property of a chart axis when the property is not directly exposed, ensuring compatibility across Aspose.Cells versions.
// Common Searches: how to read chart tick label rotation angle with Aspose.Cells C# | set Excel chart category axis labels to horizontal using Aspose.Cells | Aspose.Cells change tick label direction programmatically | use reflection to access TickLabelRotationAngle in Aspose.Cells chart | log current tick label angle before modifying chart in .NET
// Tags: Aspose.Cells modify category axis rotation | C# reflection for Aspose.Cells chart properties | horizontal Excel chart tick labels .NET | Aspose.Cells chart axis label rotation handling | set tick label direction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, obtains the first chart's category axis, uses reflection to read the current TickLabelRotationAngle, logs the angle, sets the rotation to 0 (horizontal), and saves the updated file.
class Program
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

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Adjust tick label rotation using reflection (compatible with various Aspose.Cells versions)
            Axis categoryAxis = chart.CategoryAxis;
            if (categoryAxis != null)
            {
                try
                {
                    var prop = categoryAxis.GetType().GetProperty("TickLabelRotationAngle");
                    if (prop != null && prop.CanRead && prop.CanWrite)
                    {
                        int currentAngle = (int)prop.GetValue(categoryAxis);
                        Console.WriteLine($"Current tick label rotation angle: {currentAngle} degrees");
                        prop.SetValue(categoryAxis, 0);
                        Console.WriteLine("Tick label rotation angle set to 0 (horizontal).");
                    }
                    else
                    {
                        Console.WriteLine("TickLabelRotationAngle property is not available in this version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error adjusting tick label rotation: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Unable to access CategoryAxis for rotation adjustment.");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
