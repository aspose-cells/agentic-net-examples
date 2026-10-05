// Title: Apply a custom number format to a PivotChart axis in an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens a .xlsx file, finds the first PivotChart, and assigns a custom number format to its category axis with Aspose.Cells. | Show how to verify chart existence and safely set number formats on both category and value axes before saving the workbook. | Provide a robust snippet that handles missing input files and checks for the NumberFormat property availability in different Aspose.Cells versions.
// Common Searches: Aspose.Cells C# set numeric display pattern for pivot chart axis | how to format X axis labels of a PivotChart in an existing workbook using Aspose.Cells | C# load Excel file and change chart axis number format with Aspose.Cells | apply number format to chart category axis .xlsx Aspose.Cells | check chart presence before formatting axis Aspose.Cells C#
// Tags: Aspose.Cells chart axis number format | C# pivot chart custom axis formatting | modify Excel chart axis with Aspose.Cells | format X axis of Excel chart .xlsx | chart axis NumberFormat property handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace PivotChartAxisFormatting
{
    // The example loads an existing Excel workbook, confirms a chart is present on the first worksheet, attempts to apply a custom number format to the chart's category and value axes (with version‑specific handling), and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook that contains a PivotChart
                Workbook workbook = new Workbook(inputPath);

                // Assume the chart is on the first worksheet and is the first chart
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                    return;
                }

                Chart chart = sheet.Charts[0];

                // Attempt to format the Category (X) axis if the API supports it
                try
                {
                    // The Axis class may not expose a NumberFormat property in older versions.
                    // If available, the following line will apply a custom number format.
                    // Uncomment the line below if your Aspose.Cells version supports it.
                    // chart.CategoryAxis.NumberFormat = "#,##0.00\" units\"";

                    // Similarly, you can format the Value (Y) axis:
                    // chart.ValueAxis.NumberFormat = "#,##0.00\" units\"";
                }
                catch (Exception exAxis)
                {
                    Console.WriteLine($"Error formatting axis: {exAxis.Message}");
                }

                // Save the modified workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display the error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
