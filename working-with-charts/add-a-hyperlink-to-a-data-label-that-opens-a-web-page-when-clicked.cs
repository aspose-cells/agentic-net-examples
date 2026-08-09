// Title: Add Clickable Hyperlink to Chart Data Labels with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create an Excel workbook, populate category, value and URL cells, attach Hyperlink objects, build a column chart, enable data labels, link those labels to URL cells via the LinkedSource property, and save the file so that clicking a data label opens the associated web page.
// Keywords: Aspose.Cells | C# | .NET | Excel chart hyperlink | data label link | LinkedSource property | column chart | interactive Excel workbook | hyperlink API | chart automation
// Common Searches: how to add hyperlink to chart data label using Aspose.Cells | Aspose.Cells set data label to open web page | C# link Excel chart label to URL | Aspose.Cells chart data label hyperlink example | make chart labels clickable in .NET
// Developer Intent: Generate an Excel chart whose data labels act as clickable links that open specified URLs when selected.
// Use Cases: Monthly sales chart where each column label opens a detailed dashboard for that month. | Product performance sheet that navigates to product pages from chart labels. | Executive presentation with chart labels that link to supporting documentation.
// AI Prompts: Show C# code to attach hyperlinks to chart data labels with Aspose.Cells. | Explain how to use the LinkedSource property to make chart labels open web pages. | Provide step‑by‑step instructions for creating a column chart with clickable data labels in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create an Excel workbook, populate category, value and URL cells, attach Hyperlink objects, build a column chart, enable data labels, link those labels to URL cells via the LinkedSource property, and save the file so that clicking a data label opens the associated web page.
    public class DataLabelHyperlinkDemo
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category names
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Values for the series
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Cells that will hold the hyperlink text (linked source for data labels)
            sheet.Cells["C2"].PutValue("https://example.com/jan");
            sheet.Cells["C3"].PutValue("https://example.com/feb");
            sheet.Cells["C4"].PutValue("https://example.com/mar");

            // Add hyperlinks to the cells in column C
            // These hyperlinks will be opened when the data label is clicked
            sheet.Hyperlinks.Add("C2", 1, 1, "https://example.com/jan");
            sheet.Hyperlinks.Add("C3", 1, 1, "https://example.com/feb");
            sheet.Hyperlinks.Add("C4", 1, 1, "https://example.com/mar");

            // Create a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series (values from B2:B4) and set category axis (A2:A4)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.ShowCategoryName = true;        // Show the category name
            series.DataLabels.ShowSeriesName = false;

            // Link the data labels to the cells that contain hyperlinks (C2:C4)
            series.DataLabels.LinkedSource = "C2:C4";

            // Save the workbook to an XLSX file
            workbook.Save("DataLabelHyperlinkDemo.xlsx");
        }
    }
}
