// Title: Set major tick marks on the X‑axis of the first chart in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook, accesses the first worksheet's first chart, sets the category axis major tick mark to Inside, and saves the modified file. | Show a C# example that verifies a workbook contains charts, applies inside major tick marks to the X‑axis of a chart with Aspose.Cells, and includes error handling for missing files or charts.
// Common Searches: Aspose.Cells C# set X axis major tick mark inside for first chart | how to change category axis tick marks in Excel using Aspose.Cells .NET | C# code to add major tick marks to chart axis with Aspose.Cells | example of modifying chart axis tick style in an existing workbook using Aspose.Cells | Aspose.Cells chart formatting tick marks category axis C#
// Tags: Aspose.Cells chart category axis major tick mark | C# set X axis tick mark inside Aspose.Cells | modify Excel chart axis formatting Aspose.Cells | Aspose.Cells workbook chart customization C# | Excel chart tick mark configuration .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook, checks for at least one chart on the first worksheet, sets the major tick mark of the chart's category (X) axis to Inside, and saves the updated workbook, with handling for missing files or absent charts.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
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

                // Check for at least one chart
                if (sheet.Charts.Count > 0)
                {
                    // Get the first chart
                    Chart chart = sheet.Charts[0];

                    // Set major tick marks on the X (category) axis
                    chart.CategoryAxis.MajorTickMark = TickMarkType.Inside;
                }
                else
                {
                    Console.WriteLine("No charts found in the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
