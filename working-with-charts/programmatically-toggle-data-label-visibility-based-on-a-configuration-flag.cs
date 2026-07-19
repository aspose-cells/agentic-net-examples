// Title: Toggle Data Labels on an Aspose.Cells Chart in C# with a Configurable Flag
// Description: Shows how to create a workbook, add a column chart, and programmatically show or hide series data labels using the Series.DataLabels.ShowValue property driven by a boolean configuration flag, then save the file as XLSX.
// Keywords: Aspose.Cells | C# chart data labels | Series.DataLabels.ShowValue | toggle data labels | conditional chart labels | Excel chart label visibility | configuration flag | .NET Aspose.Cells
// Common Searches: C# hide chart data labels Aspose.Cells | show data labels based on setting Aspose.Cells | conditional Series.DataLabels.ShowValue .NET | toggle Excel chart labels programmatically | use config flag to control Aspose.Cells chart labels
// Developer Intent: Control the visibility of chart data labels at runtime using a boolean setting.
// Use Cases: User‑driven report where labels appear only when a preference is enabled. | Separate visual styles for summary sheets (labels hidden) versus detailed analysis (labels shown). | Performance‑optimized export of large charts by disabling labels when not needed.
// AI Prompts: Generate C# code that reads a boolean from appsettings.json and applies it to Series.DataLabels.ShowValue in an Aspose.Cells chart. | Explain how to apply the ShowValue toggle to all series in a multi‑series chart using Aspose.Cells. | Create a reusable method that accepts a flag and sets data label visibility for any Aspose.Cells chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a column chart, and programmatically show or hide series data labels using the Series.DataLabels.ShowValue property driven by a boolean configuration flag, then save the file as XLSX.
    public class ToggleDataLabelVisibility
    {
        // Configuration flag to control data label visibility
        private static bool _showDataLabels = true; // Set to false to hide labels

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series of the chart
                Series series = chart.NSeries[0];

                // Enable data labels and set their visibility based on the configuration flag
                series.DataLabels.ShowValue = _showDataLabels;

                // Save the workbook to a file
                string outputPath = "ToggleDataLabelVisibility.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ToggleDataLabelVisibility.Run();
        }
    }
}
