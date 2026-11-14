// Title: Add a clickable timeline marker shape that links to another worksheet and preserves the hyperlink in a PDF using Aspose.Cells for .NET
// Common Searches: how to create a shape hyperlink to another worksheet in Aspose.Cells C# PDF export | Aspose.Cells timeline marker with hyperlink to details sheet in generated PDF | C# add rectangle shape with worksheet link that works in PDF using Aspose.Cells | export workbook to PDF with interactive shape navigation in Aspose.Cells .NET | link a shape to a specific cell in another sheet when saving as PDF with Aspose.Cells
// Tags: rectangle shape hyperlink to worksheet Aspose.Cells | timeline marker PDF link Aspose.Cells .NET | export workbook with interactive shape navigation | worksheet cell hyperlink preserved in PDF | add clickable shape in Aspose.Cells PDF output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates how to create a workbook with a Data sheet and a Details sheet, add a rectangle shape as a timeline marker on the Data sheet, assign a hyperlink that points to cell A1 of the Details sheet, and save the workbook as a PDF. The resulting PDF retains the clickable marker, allowing users to navigate directly to the target worksheet when the PDF is viewed.
class TimelineHyperlinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets: one for data, one for details (target of hyperlink)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            Worksheet detailsSheet = workbook.Worksheets.Add("Details");
            detailsSheet.Cells["A1"].PutValue("Details for selected timeline marker");

            // Populate sample data in the Data sheet
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            dataSheet.Cells["B3"].PutValue(150);
            dataSheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            dataSheet.Cells["B4"].PutValue(180);

            // Add a rectangle shape that will act as a timeline marker
            int markerRow = 5;      // zero‑based index (row 6 in Excel)
            int markerColumn = 0;   // column A (zero‑based)
            int markerWidth = 100;  // width in pixels
            int markerHeight = 30;  // height in pixels

            // Add the shape to the Data sheet
            Shape timelineMarker = dataSheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                markerRow,          // upper left row
                markerColumn,      // upper left column
                0,                 // top offset (pixels)
                0,                 // left offset (pixels)
                markerHeight,      // height (pixels)
                markerWidth);      // width (pixels)

            timelineMarker.Text = "Jan‑Mar 2023";

            // Create a worksheet hyperlink that points to cell A1 of the Details sheet
            string hyperlinkAddress = $"{detailsSheet.Name}!A1";
            timelineMarker.AddHyperlink(hyperlinkAddress);

            // Optional styling
            timelineMarker.Line.Weight = 1.0f;
            timelineMarker.Line.DashStyle = MsoLineDashStyle.Solid;

            // Save the workbook as PDF; the hyperlink will be preserved
            string outputPath = "TimelineWithHyperlink.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
