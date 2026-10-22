// Title: Set the legend font to Arial 12 pt for the first chart in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file, accesses the first chart on the first worksheet, and changes the legend font to Arial with a size of 12 points using Aspose.Cells. | Show how to modify a chart's Legend.Font properties (Name and Size) in Aspose.Cells for .NET. | Provide a complete example that loads a workbook, updates the legend font of the first chart, and saves the file.
// Common Searches: how to change chart legend font to Arial using Aspose.Cells C# | Aspose.Cells set legend font size 12 points in Excel chart | C# programmatically modify legend font name and size of first chart in workbook | example of accessing chart legend properties with Aspose.Cells for .NET | update Excel chart legend typography with Aspose.Cells library
// Tags: Aspose.Cells chart legend font size | Aspose.Cells set legend font name | C# modify Excel chart legend | Excel chart legend formatting Aspose.Cells | first worksheet chart legend customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, accesses the first worksheet's first chart, sets the legend font to Arial 12 pt, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load an existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Access the first chart
                Chart chart = sheet.Charts[0];

                // Change the legend font size to 12 points
                chart.Legend.Font.Size = 12;

                // Set the legend font name to Arial
                chart.Legend.Font.Name = "Arial";
            }
            else
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
