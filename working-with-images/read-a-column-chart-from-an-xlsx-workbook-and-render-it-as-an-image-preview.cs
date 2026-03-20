using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartPreviewExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that contains the chart
            string sourceFile = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourceFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all charts in the worksheet
            foreach (Chart chart in worksheet.Charts)
            {
                // Check if the chart is a column chart
                if (chart.Type == ChartType.Column)
                {
                    // Define the output image file name (PNG format)
                    string imageFile = "ChartPreview.png";

                    // Render the chart to an image file using the ToImage method
                    chart.ToImage(imageFile, ImageType.Png);

                    Console.WriteLine($"Column chart exported to image: {imageFile}");
                }
            }
        }
    }
}