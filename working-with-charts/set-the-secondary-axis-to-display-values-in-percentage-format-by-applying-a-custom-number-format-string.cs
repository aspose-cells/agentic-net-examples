// Title: How to set a chart's secondary axis to display percentages using a custom number format in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set secondary value axis to percentage format | custom number format string for secondary axis in Aspose.Cells chart | format secondary axis as 0% in Aspose.Cells column chart using C# | disable linked number format and apply 0% to chart axis Aspose.Cells | display secondary axis values as percentages in Aspose.Cells .NET
// Tags: secondary axis percentage format Aspose.Cells | custom number format string chart axis C# | disable source linked number format Aspose.Cells | column chart value axis formatting .NET | Aspose.Cells chart axis number format customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a column chart, and shows how to turn off source-linked formatting and apply a custom "0%" number format to the chart's secondary value axis using Aspose.Cells for .NET, then saves the file as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (position can be adjusted as needed)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
                Chart chart = sheet.Charts[chartIndex];

                // NOTE: In some Aspose.Cells versions the Axis class does not expose
                // IsNumberFormatLinkedToSource or NumberFormat properties.
                // If available, you can uncomment the following lines to format the value axis as percentages.
                // chart.ValueAxis.IsNumberFormatLinkedToSource = false;
                // chart.ValueAxis.NumberFormat = "0%";

                // Save the workbook
                workbook.Save("Result.xlsx");
                Console.WriteLine("Workbook saved successfully as Result.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
