using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Set the regional settings for the workbook (e.g., Japanese locale)
            workbook.Settings.Region = CountryCode.Japan;

            // 3. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 4. Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1500);
            sheet.Cells["B3"].PutValue(2300);
            sheet.Cells["B4"].PutValue(1800);

            // 5. Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // 6. Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // 7. Convert the chart to an image (PNG format) and save it to a file
            chart.ToImage("LocalizedChart.png", ImageType.Png);

            // 8. Optionally, save the workbook to verify the regional settings are applied
            workbook.Save("LocalizedChartWorkbook.xlsx");

            Console.WriteLine("Chart rendered with localized region and saved as image.");
        }
    }
}