// Title: Export a column chart from an Aspose.Cells workbook to an SVG file with selectable text in C#
// AI Prompts: Generate C# code that creates a column chart in an Aspose.Cells workbook and uses ImageOrPrintOptions to save the chart as an SVG file with selectable text. | Show how to export only a chosen chart from a worksheet to SVG using Aspose.Cells' Chart.ToImage method and configure the output for web‑ready scalability. | Provide a C# example that iterates over all charts on a worksheet and writes each one to a separate SVG file while preserving editable text.
// Common Searches: how to save an Aspose.Cells chart as an SVG with editable text in C# | C# Aspose.Cells export specific chart to SVG using ImageOrPrintOptions | export multiple Excel charts to separate SVG files with selectable text using Aspose.Cells | Aspose.Cells Chart.ToImage SVG example C#
// Tags: Aspose.Cells export chart to SVG | C# ImageOrPrintOptions SVG chart output | selectable text in SVG from Excel chart | Chart.ToImage method Aspose.Cells | multiple chart SVG export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering; // Required for ImageOrPrintOptions

// Creates a workbook, adds sample data, builds a column chart, sets its title, configures ImageOrPrintOptions with ImageType.Svg, and uses Chart.ToImage to write the chart to an SVG file where text remains selectable; optionally saves the workbook as an XLSX file.
class ExportChartToSvg
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one with Workbook workbook = new Workbook("input.xlsx");)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a title for the chart
            chart.Title.Text = "Quarterly Sales";

            // Configure image options for SVG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Svg // Export as SVG
            };

            // Export the chart directly to an SVG file
            string svgPath = "ChartOutput.svg";
            chart.ToImage(svgPath, imgOptions);

            // (Optional) Save the workbook if you need the Excel file as well
            string workbookPath = "WorkbookWithChart.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
