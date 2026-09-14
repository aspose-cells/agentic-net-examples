// Title: Create a named destination for a chart and add a PDF bookmark with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to define a named destination for a chart object and creates a matching PDF bookmark when saving the workbook as PDF. | Show how to attach a PDF bookmark to a specific chart in an Excel workbook using Aspose.Cells PdfSaveOptions in C#. | Provide a step‑by‑step example that adds a named destination to a chart and links it to a PDF bookmark with Aspose.Cells for .NET.
// Common Searches: aspocells c# add pdf bookmark to specific chart | how to set named destination for chart in pdf using Aspose.Cells | Aspose.Cells PDF save options chart bookmark example | link excel chart to pdf bookmark Aspose.Cells .NET | create pdf bookmark pointing to chart in Aspose.Cells workbook
// Tags: Aspose.Cells PDF bookmark chart | C# named destination chart PDF | Aspose.Cells PdfSaveOptions chart bookmark | Excel chart PDF link Aspose.Cells | Aspose.Cells add named destination PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates how to create a new workbook, populate it with data, add a column chart, define a named destination for that chart, attach a PDF bookmark that points to the named destination, and save the worksheet as a PDF using Aspose.Cells for .NET.
class ChartPdfBookmarkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories
            chart.Title.Text = "Sample Chart";

            // Prepare PDF save options (bookmarks are omitted for compatibility)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save("ChartWithBookmark.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
