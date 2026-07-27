using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class SaveChartsAsJpeg
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ----- First worksheet with a column chart -----
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Cells["A1"].PutValue("Category");
        ws1.Cells["A2"].PutValue("Apple");
        ws1.Cells["A3"].PutValue("Orange");
        ws1.Cells["A4"].PutValue("Banana");
        ws1.Cells["B1"].PutValue("Value");
        ws1.Cells["B2"].PutValue(10);
        ws1.Cells["B3"].PutValue(15);
        ws1.Cells["B4"].PutValue(7);

        int chartIdx1 = ws1.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart1 = ws1.Charts[chartIdx1];
        chart1.NSeries.Add("B2:B4", true);
        chart1.NSeries.CategoryData = "A2:A4";

        // ----- Second worksheet with a pie chart (optional) -----
        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
        ws2.Cells["A1"].PutValue("Category");
        ws2.Cells["A2"].PutValue("X");
        ws2.Cells["A3"].PutValue("Y");
        ws2.Cells["B1"].PutValue("Value");
        ws2.Cells["B2"].PutValue(20);
        ws2.Cells["B3"].PutValue(30);

        int chartIdx2 = ws2.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart2 = ws2.Charts[chartIdx2];
        chart2.NSeries.Add("B2:B3", true);
        chart2.NSeries.CategoryData = "A2:A3";

        // ----- Save each chart as a JPEG named after its worksheet -----
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            for (int i = 0; i < sheet.Charts.Count; i++)
            {
                Chart c = sheet.Charts[i];

                // Build file name: WorksheetName.jpg (or WorksheetName_ChartN.jpg if multiple charts)
                string fileName = sheet.Name;
                if (sheet.Charts.Count > 1)
                {
                    fileName += $"_Chart{i + 1}";
                }
                fileName += ".jpg";

                // Save chart image as JPEG
                c.ToImage(fileName, ImageType.Jpeg);
            }
        }

        // Optionally save the workbook itself
        workbook.Save("ChartsWorkbook.xlsx");
    }
}