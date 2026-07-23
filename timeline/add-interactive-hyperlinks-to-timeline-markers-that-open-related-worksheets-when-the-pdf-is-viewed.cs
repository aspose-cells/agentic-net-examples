// Title: Add internal worksheet hyperlinks to an Aspose.Cells Timeline and keep them clickable in the PDF (C#)
// Description: Demonstrates how to create a workbook with a Data sheet, add a pivot table, insert a Timeline control linked to the Date field, assign an internal hyperlink to the Timeline shape that points to a cell in a Details worksheet, and export the workbook to PDF while preserving the clickable link.
// Keywords: Aspose.Cells Timeline hyperlink | C# timeline PDF link | internal worksheet hyperlink Aspose.Cells | export timeline to PDF | clickable timeline marker | Aspose.Cells PDF navigation | timeline shape hyperlink | Aspose.Cells C# example
// Common Searches: how to add hyperlink to Aspose.Cells timeline | timeline control link to another worksheet PDF | Aspose.Cells C# export timeline with clickable links | internal cell hyperlink in PDF using Aspose.Cells | navigate from timeline marker to worksheet in PDF
// Developer Intent: Create a Timeline control whose shape contains an internal hyperlink that opens a specific worksheet when the generated PDF is viewed.
// Use Cases: Drill‑down from a summary timeline to a detailed data sheet in a PDF report. | Provide one‑click navigation from timeline markers to supporting worksheets for auditors or analysts. | Enhance PDF interactivity by linking timeline events to related documentation within the same workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a unique hyperlink to each timeline marker based on its date range, pointing to corresponding worksheets. | Show how to customize the display text and address of a Timeline shape hyperlink before saving to PDF. | Explain steps to verify that timeline hyperlinks remain functional after converting an Excel workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook with a Data sheet, add a pivot table, insert a Timeline control linked to the Date field, assign an internal hyperlink to the Timeline shape that points to a cell in a Details worksheet, and export the workbook to PDF while preserving the clickable link.
class TimelineHyperlinkDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it to "Data"
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Add a second worksheet named "Details"
            Worksheet detailsSheet = workbook.Worksheets.Add("Details");

            // Populate the "Data" sheet with sample date/value data
            Cells dataCells = dataSheet.Cells;
            dataCells["A1"].Value = "Date";
            dataCells["B1"].Value = "Value";
            dataCells["A2"].Value = new DateTime(2023, 1, 1);
            dataCells["B2"].Value = 100;
            dataCells["A3"].Value = new DateTime(2023, 2, 1);
            dataCells["B3"].Value = 200;

            // Create a pivot table based on the sample data
            int pivotIndex = dataSheet.PivotTables.Add("A1:B3", "D1", "Pivot1");
            PivotTable pivot = dataSheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the pivot table (using the "Date" field)
            int timelineIndex = dataSheet.Timelines.Add(pivot, "E1", "Date");
            Timeline timeline = dataSheet.Timelines[timelineIndex];

            // Obtain the shape representing the Timeline
            TimelineShape timelineShape = timeline.Shape;

            // Set an internal hyperlink that points to cell A1 of the "Details" worksheet
            timelineShape.Hyperlink.Address = "Details!A1";          // Internal cell reference
            timelineShape.Hyperlink.TextToDisplay = "Go to Details"; // Display text

            // Save the workbook as a PDF; the Timeline shape will retain the hyperlink in the PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("TimelineHyperlink.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
