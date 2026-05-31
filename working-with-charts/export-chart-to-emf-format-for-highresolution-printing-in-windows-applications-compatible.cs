using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class ExportChartToEmf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Cherry");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure image options for high‑resolution EMF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            // Set the image type to EMF
            ImageType = ImageType.Emf,

            // High resolution (e.g., 300 DPI)
            HorizontalResolution = 300,
            VerticalResolution = 300,

            // Prefer EMF+ records when rendering (optional)
            EmfRenderSetting = EmfRenderSetting.EmfPlusPrefer
        };

        // Export the chart to an EMF file using the options
        chart.ToImage("ChartOutput.emf", options);

        // (Optional) Save the workbook if you need the Excel file as well
        workbook.Save("ChartWorkbook.xlsx");
    }
}