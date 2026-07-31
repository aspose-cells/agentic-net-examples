// Title: Export Aspose.Cells Chart to High‑Resolution EMF (300 DPI) in C#
// Description: Demonstrates how to create a workbook, add a column chart, configure ImageOrPrintOptions for 300 DPI EMF output with EmfOnly rendering, and export the chart to a vector EMF file suitable for Windows printing and embedding in Office documents.
// Keywords: Aspose.Cells | C# | chart export | EMF | high resolution | 300 DPI | ImageOrPrintOptions | EmfRenderSetting | vector graphics | Windows printing | Excel chart to EMF
// Common Searches: Aspose.Cells export chart to EMF C# | C# generate high DPI EMF from Excel chart | how to save Aspose.Cells chart as EMF | 300 DPI EMF export Aspose.Cells | EMF vector chart for Windows printing C#
// Developer Intent: Create a 300 DPI EMF vector image of an Excel chart using Aspose.Cells in C#.
// Use Cases: Produce printable, scalable charts for Windows‑based reports. | Embed high‑quality vector charts in Word or PowerPoint presentations. | Automate batch export of workbook charts to EMF files for a reporting pipeline. | Generate graphics for UI components that require resolution‑independent rendering.
// AI Prompts: Write C# code to export a pie chart from an Aspose.Cells workbook to a 600 DPI EMF file. | Explain how to include chart background and borders when exporting to EMF with Aspose.Cells. | Provide a script that loops through all worksheets and saves each chart as a separate EMF file. | Show how to combine multiple EMF chart files into a single PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartToEmf
{
    // Demonstrates how to create a workbook, add a column chart, configure ImageOrPrintOptions for 300 DPI EMF output with EmfOnly rendering, and export the chart to a vector EMF file suitable for Windows printing and embedding in Office documents.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image options for high‑resolution EMF output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Emf,                 // Export as EMF
                HorizontalResolution = 300,                // 300 DPI horizontal
                VerticalResolution = 300,                  // 300 DPI vertical
                EmfRenderSetting = EmfRenderSetting.EmfOnly // Use EMF records only
            };

            // Export the chart to an EMF file using the options above
            chart.ToImage("ChartOutput.emf", imgOptions);

            // Optionally save the workbook (not required for EMF export but follows lifecycle rule)
            workbook.Save("ChartWorkbook.xlsx");

            Console.WriteLine("Chart successfully exported to EMF format with high resolution.");
        }
    }
}
