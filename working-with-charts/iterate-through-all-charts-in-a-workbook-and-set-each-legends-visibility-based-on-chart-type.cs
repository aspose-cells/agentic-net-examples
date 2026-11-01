// Title: Hide legends for pie and doughnut charts while showing them for other chart types in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide legend for pie chart in an existing workbook | How to set ShowLegend property for each chart based on type with Aspose.Cells | Iterate through all worksheets and charts to toggle legend visibility in .NET | C# code to hide doughnut chart legend using Aspose.Cells API | Loop over Excel charts and change legend visibility according to chart type
// Tags: hide chart legend Aspose.Cells C# | iterate workbook charts Aspose.Cells | conditional legend visibility based on chart type | pie and doughnut chart legend handling Aspose.Cells | showlegend property manipulation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel file, loops through every worksheet and each chart, disables the legend for Pie and Doughnut charts while enabling it for all other chart types, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each chart on the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Hide legend for Pie and Doughnut charts; show for others
                        if (chart.Type == ChartType.Pie || chart.Type == ChartType.Doughnut)
                        {
                            chart.ShowLegend = false; // Hide legend
                        }
                        else
                        {
                            chart.ShowLegend = true;  // Show legend
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
