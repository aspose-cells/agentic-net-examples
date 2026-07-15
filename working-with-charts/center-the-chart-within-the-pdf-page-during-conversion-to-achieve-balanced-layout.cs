using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCenterPdf
{
    public class CenterChartInPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Define PDF page size (standard Letter size) in inches
                float pageWidth = 8.5f;
                float pageHeight = 11f;

                // Export the chart to PDF, centering it horizontally and vertically
                chart.ToPdf("CenteredChart.pdf", pageWidth, pageHeight,
                            PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);

                Console.WriteLine("Chart has been exported to PDF with centered alignment.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            CenterChartInPdf.Run();
        }
    }
}