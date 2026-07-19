// Title: Add Hyperlink to Chart Data Labels via Linked Cells – Aspose.Cells for .NET (C#)
// Description: Creates a workbook with a column chart, links data labels to cells, assigns hyperlinks to those cells, and makes each label open a web page when clicked.
// Keywords: Aspose.Cells chart hyperlink C# | hyperlink data label Aspose.Cells | linked source data label Excel | add clickable label to chart | Aspose.Cells .NET chart example
// Common Searches: Aspose.Cells add hyperlink to chart label | link chart data label to URL in C# | hyperlinked data labels Excel Aspose | chart label click opens web page Aspose.Cells | C# example for chart data label hyperlink
// Developer Intent: Enable chart data labels to act as clickable links that open specific URLs.
// Use Cases: Sales dashboard where each column label links to the product detail page. | Financial report with chart labels that navigate to supporting documents. | Interactive Excel export for presentations, allowing quick access to external resources via label clicks.
// AI Prompts: Generate C# code using Aspose.Cells to create a column chart with data labels that open different URLs. | Explain how to link chart data labels to cells and attach hyperlinks in Aspose.Cells for .NET. | Show step‑by‑step how to set TextToDisplay for hyperlinked chart data labels in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with a column chart, links data labels to cells, assigns hyperlinks to those cells, and makes each label open a web page when clicked.
class AddHyperlinkToDataLabel
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Cells that will hold the display text for data labels
        // (these cells will be linked to the data labels)
        sheet.Cells["C1"].PutValue("Label");
        sheet.Cells["C2"].PutValue("Link A");
        sheet.Cells["C3"].PutValue("Link B");
        sheet.Cells["C4"].PutValue("Link C");

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and link them to the cells in column C
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;               // show the value
        series.DataLabels.LinkedSource = "C2:C4";          // link label text to cells C2:C4
        series.DataLabels.NumberFormatLinked = false;     // keep number format independent

        // Add hyperlinks to the cells that are linked to the data labels
        // When a user clicks a data label, the hyperlink of the linked cell will be followed.
        sheet.Hyperlinks.Add("C2", 1, 1, "https://example.com/pageA");
        sheet.Hyperlinks.Add("C3", 1, 1, "https://example.com/pageB");
        sheet.Hyperlinks.Add("C4", 1, 1, "https://example.com/pageC");

        // Optionally set display text for the hyperlinks (the text shown in the data label)
        sheet.Hyperlinks[0].TextToDisplay = "Page A";
        sheet.Hyperlinks[1].TextToDisplay = "Page B";
        sheet.Hyperlinks[2].TextToDisplay = "Page C";

        // Save the workbook
        workbook.Save("ChartWithHyperlinkedDataLabels.xlsx");
    }
}
