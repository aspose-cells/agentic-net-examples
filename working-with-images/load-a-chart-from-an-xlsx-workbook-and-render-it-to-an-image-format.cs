using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RenderChartFromWorkbook
{
    static void Main()
    {
        // Path to the Excel file that contains the chart
        string sourcePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one chart
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the workbook.");
            return;
        }

        // Retrieve the first chart in the worksheet
        Chart chart = sheet.Charts[0];

        // Define the output image file; the format is taken from the file extension (PNG here)
        string imagePath = "chart_output.png";

        // Render the chart to the image file
        chart.ToImage(imagePath);

        Console.WriteLine($"Chart successfully rendered to: {imagePath}");
    }
}