// Title: C# – Set Chart Background to Light Gray and Remove Fill Patterns with Aspose.Cells
// Description: Creates a new workbook, adds a column chart, sets both the ChartArea and PlotArea background to LightGray, clears any fill pattern using FillPattern.None, and saves the file as an XLSX workbook.
// Keywords: Aspose.Cells | C# | chart background color | light gray chart area | remove fill pattern | ChartArea formatting | PlotArea formatting | FillPattern.None | Excel chart styling | .NET chart example
// Common Searches: Aspose.Cells change chart background color C# | remove chart fill pattern Aspose.Cells .NET | set light gray background for chart area Aspose | clear fill pattern in chart area using Aspose.Cells | Aspose.Cells chart area formatting example
// Developer Intent: Apply a solid light‑gray background to a chart and eliminate any default fill patterns.
// Use Cases: Standardize the appearance of all charts in a financial report with a neutral light‑gray background. | Prepare printable Excel dashboards where pattern fills could interfere with scanning or printing. | Implement corporate branding by ensuring chart backgrounds match the company’s color palette without textures.
// AI Prompts: Generate C# code with Aspose.Cells to set a chart’s background to a specific RGB value and remove its fill pattern. | Show how to apply the same background and fill‑pattern settings to multiple charts in a workbook using Aspose.Cells for .NET. | Explain how to revert chart area formatting to defaults after customizing the background color in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackground
{
    // Creates a new workbook, adds a column chart, sets both the ChartArea and PlotArea background to LightGray, clears any fill pattern using FillPattern.None, and saves the file as an XLSX workbook.
    public class SetChartBackground
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart area background color to light gray
            chart.ChartArea.Area.BackgroundColor = Color.LightGray;
            // Remove any fill pattern from the chart area
            chart.ChartArea.Area.FillFormat.Pattern = FillPattern.None;

            // Clear the plot area fill pattern and set its background color
            chart.PlotArea.Area.BackgroundColor = Color.LightGray;
            chart.PlotArea.Area.FillFormat.Pattern = FillPattern.None;

            // Save the workbook
            string outputPath = "ChartWithLightGrayBackground.xlsx";
            workbook.Save(outputPath);
        }
    }
}
