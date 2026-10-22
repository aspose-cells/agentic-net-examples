// Title: Export an Aspose.Cells chart to PDF using a MemoryStream in C#
// AI Prompts: Write C# code that creates a worksheet chart and saves the workbook as a PDF directly into a MemoryStream with Aspose.Cells. | Show how to reset a MemoryStream after saving a chart PDF with Aspose.Cells so the data can be read or transmitted elsewhere. | Demonstrate extracting the PDF byte array from a MemoryStream after exporting a chart with Aspose.Cells without writing a file.
// Common Searches: how to export a chart from Aspose.Cells to a PDF stream in C# | Aspose.Cells save workbook as PDF to MemoryStream without creating a file | C# in‑memory PDF generation for Excel chart using Aspose.Cells | resetting MemoryStream position after Aspose.Cells PDF save for further processing
// Tags: Aspose.Cells chart PDF memory stream | C# save workbook as PDF to stream | in‑memory PDF export Aspose.Cells | MemoryStream position reset after PDF export | Aspose.Cells SaveFormat.Pdf stream output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data and a column chart, then saves the workbook as a PDF directly into a MemoryStream, resets the stream position, and leaves the PDF bytes available for further in‑memory processing.
class ChartPdfToMemoryStream
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Optional: set chart title
        chart.Title.Text = "Sample Column Chart";

        // Prepare a memory stream to hold the PDF output
        using (MemoryStream pdfStream = new MemoryStream())
        {
            // Save the workbook (including the chart) as PDF into the memory stream
            workbook.Save(pdfStream, SaveFormat.Pdf);

            // Reset the stream position to the beginning for further processing
            pdfStream.Position = 0;

            // Example: write the PDF bytes to a file (optional, can be omitted)
            // File.WriteAllBytes("ChartOutput.pdf", pdfStream.ToArray());

            // At this point, pdfStream contains the PDF data and can be used for further in‑memory processing
        }
    }
}
