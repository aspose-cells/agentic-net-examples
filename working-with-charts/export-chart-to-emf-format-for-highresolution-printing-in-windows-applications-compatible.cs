// Title: Export a Workbook Chart to 300 DPI EMF Using Aspose.Cells for .NET
// Description: Creates a sample workbook, adds a column chart, configures ImageOrPrintOptions for 300 DPI vector output and EMF+ preference, then saves the chart as an EMF file while optionally keeping the workbook.
// Keywords: Aspose.Cells EMF export | C# chart to EMF | ImageOrPrintOptions 300 DPI | EMF+ Aspose.Cells | export column chart .NET | vector chart printing | Aspose.Cells chart image options
// Common Searches: Aspose.Cells export chart EMF C# | 300 DPI EMF chart Aspose.Cells | set EMF+ rendering in Aspose.Cells | export all workbook charts as EMF using Aspose.Cells | save chart as vector image for Windows reports
// Developer Intent: Generate a 300 DPI EMF vector image of a worksheet chart in C#.
// Use Cases: Insert a scalable chart into Windows reporting tools that require vector graphics. | Create printable graphics for inclusion in Microsoft Office documents. | Batch‑export multiple charts from a workbook for design assets. | Maintain the original workbook while providing a separate high‑quality EMF file.
// AI Prompts: Show how to change the exported EMF size and add a transparent background with Aspose.Cells. | Provide code that loops through every chart in a workbook and saves each as a 600 DPI EMF file. | Explain how to embed the generated EMF chart into a Word document using Aspose.Words after export. | Suggest ways to improve rendering speed when exporting many charts to EMF in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a sample workbook, adds a column chart, configures ImageOrPrintOptions for 300 DPI vector output and EMF+ preference, then saves the chart as an EMF file while optionally keeping the workbook.
class ExportChartToEmf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories
        chart.Title.Text = "Sample Chart";

        // Configure high‑resolution image options for EMF output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Emf,                     // Export as EMF
            HorizontalResolution = 300,                    // 300 DPI horizontal
            VerticalResolution = 300,                      // 300 DPI vertical
            EmfRenderSetting = EmfRenderSetting.EmfPlusPrefer // Prefer EMF+ records (optional)
        };

        // Export the chart to an EMF file using the specified options
        chart.ToImage("ChartOutput.emf", imgOptions);

        // Save the workbook (optional, just to keep the chart in the file)
        workbook.Save("WorkbookWithChart.xlsx");
    }
}
