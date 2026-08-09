// Title: Export Aspose.Cells Chart to High‑Resolution EMF (300 DPI) in C# for Windows Printing
// Description: Learn how to create a workbook, add a column chart, configure ImageOrPrintOptions for 300 DPI EMF+ output, and save the chart as a vector EMF file suitable for high‑quality Windows desktop printing using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart export EMF | C# high resolution EMF | 300 DPI chart Aspose | EMF+ rendering Aspose.Cells | vector chart Windows printing | ImageOrPrintOptions EMF C# | export Excel chart to EMF
// Common Searches: Aspose.Cells export chart to EMF C# | how to save chart as EMF with 300 DPI | C# generate high‑resolution EMF from Excel chart | EMF+ chart export Aspose.Cells example | set DPI for chart image Aspose.Cells
// Developer Intent: Generate a high‑resolution EMF file from an Aspose.Cells chart for precise Windows printing.
// Use Cases: Embedding scalable vector charts in Windows desktop reports. | Creating print‑ready graphics for Office documents or PDFs. | Supplying DPI‑controlled chart assets to .NET applications that require vector output.
// AI Prompts: Provide C# code that exports an Aspose.Cells pie chart to EMF at 600 DPI. | Show how to force EmfRenderSetting.EmfPlusOnly when saving a chart as EMF. | Explain steps to embed an exported EMF chart into a WPF Image control.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartToEmf
{
    // Learn how to create a workbook, add a column chart, configure ImageOrPrintOptions for 300 DPI EMF+ output, and save the chart as a vector EMF file suitable for high‑quality Windows desktop printing using Aspose.Cells for .NET.
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
                EmfRenderSetting = EmfRenderSetting.EmfPlusPrefer // Prefer EMF+ records
            };

            // Export the chart to an EMF file using the options
            chart.ToImage("HighResChart.emf", imgOptions);

            // (Optional) Save the workbook if you also need the Excel file
            workbook.Save("ChartWorkbook.xlsx");

            Console.WriteLine("Chart exported to HighResChart.emf with 300 DPI.");
        }
    }
}
