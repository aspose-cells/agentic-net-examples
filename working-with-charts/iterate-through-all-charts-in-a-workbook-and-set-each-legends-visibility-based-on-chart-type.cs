// Title: C# iterate all charts in an Excel workbook and set legend visibility based on chart type with Aspose.Cells
// Common Searches: Aspose.Cells C# set chart legend visibility only for pie and doughnut charts | How to loop through every chart in a workbook and hide legends for non‑pie charts using Aspose.Cells | C# code to change ShowLegend property based on ChartType in Aspose.Cells | Iterate worksheets and charts to toggle legend display in Aspose.Cells .NET | Programmatically control chart legends per chart type with Aspose.Cells for .NET
// Tags: set chart legend visibility Aspose.Cells | iterate workbook charts C# | conditional legend display based on ChartType | pie doughnut chart legend handling Aspose.Cells | hide legends for non‑pie charts .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook (creating a placeholder if missing), iterates through each worksheet and every chart, enables the legend only for Pie and Doughnut chart types while disabling it for all other types, and then saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create an empty workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var emptyWorkbook = new Workbook();
                emptyWorkbook.Save(inputPath);
                Console.WriteLine($"Input file not found. Created empty workbook at '{inputPath}'.");
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their charts.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Show legend for Pie and Doughnut charts; hide for others.
                    if (chart.Type == ChartType.Pie || chart.Type == ChartType.Doughnut)
                    {
                        chart.ShowLegend = true;
                    }
                    else
                    {
                        chart.ShowLegend = false;
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
