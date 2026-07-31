// Title: Add Clickable Hyperlinks to Chart Data Labels with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, inserts a column chart, links the chart's data labels to a cell range, adds hyperlinks to those cells, and customizes the label text so clicking a data label opens the target web page.
// Keywords: Aspose.Cells | C# | .NET | chart data label hyperlink | add hyperlink to chart label | linked source data labels | Excel chart hyperlink | Aspose.Cells example | GitHub Aspose.Cells | hyperlink in data label | column chart Aspose.Cells
// Common Searches: Aspose.Cells add hyperlink to chart data label C# | How to make chart data labels clickable in Aspose.Cells | C# Aspose.Cells linked source with hyperlink | Aspose.Cells chart label hyperlink example | Create Excel chart with clickable labels using Aspose.Cells
// Developer Intent: Create a chart whose data labels act as clickable links to external URLs.
// Use Cases: Sales dashboard where each column label opens a detailed sales report page. | Product performance chart linking each label to its product information page. | Automated financial workbook that provides direct access to supporting documents via chart labels.
// AI Prompts: Generate C# code using Aspose.Cells to add a column chart and set data labels to linked cells containing different hyperlinks. | Show how to assign custom display text to chart data label hyperlinks in Aspose.Cells for .NET. | Explain steps to update or remove hyperlinks from linked data label cells after the chart is created.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, inserts a column chart, links the chart's data labels to a cell range, adds hyperlinks to those cells, and customizes the label text so clicking a data label opens the target web page.
    public class DataLabelHyperlinkDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["A4"].PutValue("Item 3");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels and link them to a separate range that will hold the display text
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Populate the linked source range for data labels
                sheet.Cells["C2"].PutValue("Link 1");
                sheet.Cells["C3"].PutValue("Link 2");
                sheet.Cells["C4"].PutValue("Link 3");

                // Link the data labels to the created range
                series.DataLabels.LinkedSource = "C2:C4";

                // Add hyperlinks to the cells that serve as the linked source for the data labels
                sheet.Hyperlinks.Add("C2", 1, 1, "https://example.com/page1");
                sheet.Hyperlinks.Add("C3", 1, 1, "https://example.com/page2");
                sheet.Hyperlinks.Add("C4", 1, 1, "https://example.com/page3");

                // Optionally, customize the display text of the hyperlinks (the text shown on the label)
                sheet.Hyperlinks[0].TextToDisplay = "Visit Page 1";
                sheet.Hyperlinks[1].TextToDisplay = "Visit Page 2";
                sheet.Hyperlinks[2].TextToDisplay = "Visit Page 3";

                // Determine output path and save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DataLabelHyperlinkDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
