// Title: Export Aspose.Cells Timeline to PDF in a MemoryStream (C#) – No Disk I/O
// Description: Shows how to build a workbook with a pivot table and a timeline, then save the workbook as a PDF directly into a MemoryStream using Aspose.Cells for .NET, allowing further processing without creating a physical file.
// Keywords: Aspose.Cells | timeline PDF export | MemoryStream PDF C# | Aspose.Cells pivot timeline | save PDF to stream | in‑memory PDF generation | C# Aspose.Cells export | no file I/O PDF | Aspose.Cells PDF options | export workbook to stream
// Common Searches: Aspose.Cells export timeline to PDF stream C# | Save workbook with timeline to MemoryStream Aspose.Cells | Generate PDF from timeline without writing file | Aspose.Cells PDF in‑memory export example | C# export pivot timeline to PDF stream
// Developer Intent: Create a PDF that includes a timeline and write it directly to a MemoryStream for downstream use.
// Use Cases: Attach a sales‑timeline PDF to an email without creating a temporary file. | Upload the generated PDF stream to a REST API or cloud storage service directly from memory. | Persist PDF bytes in a database BLOB after in‑memory generation.
// AI Prompts: Provide C# code that adds a timeline to a pivot table and saves the workbook as a PDF into a MemoryStream using Aspose.Cells. | Explain how to reset the MemoryStream position and retrieve the PDF byte array for uploading to a web service. | Show how to configure PDF export options (page size, orientation, compression) when exporting a workbook with a timeline to a stream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Shows how to build a workbook with a pivot table and a timeline, then save the workbook as a PDF directly into a MemoryStream using Aspose.Cells for .NET, allowing further processing without creating a physical file.
class ExportTimelineToPdfStream
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate worksheet with sample date and sales data
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-3));
            worksheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-2));
            worksheet.Cells["A4"].PutValue(DateTime.Now.AddDays(-1));
            worksheet.Cells["A5"].PutValue(DateTime.Now);
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(200);
            worksheet.Cells["B5"].PutValue(250);

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivot = worksheet.PivotTables[pivotIndex];

            // Add the date field to the Page area (required for timelines) and the sales field to the Data area
            pivot.AddFieldToArea(PivotFieldType.Page, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();

            // Add a timeline linked to the pivot table's date field
            int timelineIndex = worksheet.Timelines.Add(pivot, 0, 0, "Date");
            Timeline timeline = worksheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            // Export the workbook (which includes the timeline) to a PDF stored in a memory stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0; // Reset for further processing

                Console.WriteLine($"Generated PDF stream length: {pdfStream.Length} bytes");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
