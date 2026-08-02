// Title: Export Aspose.Cells Timeline to PDF via MemoryStream in C# (No Disk I/O)
// Description: Demonstrates how to build a workbook, add a pivot table and a linked timeline, then save the entire workbook as a PDF directly into a MemoryStream using Aspose.Cells for .NET, eliminating any file‑system writes.
// Keywords: Aspose.Cells timeline PDF | C# MemoryStream PDF export | save workbook to stream | pivot table timeline Aspose | in‑memory PDF generation | Aspose.Cells .NET PDF stream | export workbook without file | timeline to PDF C# | Aspose.Cells API PDF | PDF generation from worksheet
// Common Searches: Aspose.Cells export timeline to PDF in memory | C# save workbook as PDF to MemoryStream | How to generate PDF from timeline without disk | Aspose.Cells pivot timeline PDF stream example | Create PDF from Excel workbook in memory C#
// Developer Intent: Produce a PDF that includes a timeline linked to a pivot table and keep the result in a MemoryStream for further processing.
// Use Cases: Attach the PDF to an email directly from memory. | Return the PDF as an HTTP response in a web API. | Store the PDF bytes in a database or cloud blob without creating a temporary file. | Pass the PDF stream to another service for batch processing.
// AI Prompts: Show C# code that adds a timeline to a pivot table and saves the workbook as a PDF into a MemoryStream using Aspose.Cells. | Generate a byte array from an Aspose.Cells workbook containing a timeline for use in an ASP.NET Core file download. | Explain how to configure Aspose.Cells PDF options when exporting a timeline to an in‑memory stream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Demonstrates how to build a workbook, add a pivot table and a linked timeline, then save the entire workbook as a PDF directly into a MemoryStream using Aspose.Cells for .NET, eliminating any file‑system writes.
class ExportTimelineToPdfMemoryStream
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with sample date and value data
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-3));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-2));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(DateTime.Now.AddDays(-1));
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue(DateTime.Now);
            sheet.Cells["B5"].PutValue(40);

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh the pivot table to apply changes
            pivot.RefreshData();

            // Add a timeline linked to the pivot table's date field
            int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];
            timeline.Caption = "Sales Timeline";

            // Export the workbook (including the timeline) to a PDF stored in a memory stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                Console.WriteLine($"PDF generated in memory. Stream length: {pdfStream.Length} bytes");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
