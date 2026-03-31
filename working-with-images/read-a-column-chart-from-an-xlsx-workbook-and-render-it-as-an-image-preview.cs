using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source Excel file containing the chart
        string excelPath = "input.xlsx";

        // Path where the chart preview image will be saved
        string imagePath = "chart_preview.png";

        // Load the workbook from the file
        Workbook workbook = new Workbook(excelPath);

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Find the first column chart in the worksheet
        Chart columnChart = null;
        foreach (Chart chart in worksheet.Charts)
        {
            if (chart.Type == ChartType.Column)
            {
                columnChart = chart;
                break;
            }
        }

        // If a column chart is found, render it to an image file
        if (columnChart != null)
        {
            // Save the chart as a PNG image (preview)
            columnChart.ToImage(imagePath, ImageType.Png);
            Console.WriteLine($"Chart preview saved to: {imagePath}");
        }
        else
        {
            Console.WriteLine("No column chart found in the workbook.");
        }
    }
}