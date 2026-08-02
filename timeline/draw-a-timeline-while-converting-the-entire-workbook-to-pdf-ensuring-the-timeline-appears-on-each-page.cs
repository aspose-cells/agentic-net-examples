// Title: Create a Pivot Table Timeline and Export to PDF with Timeline on Every Page – Aspose.Cells for .NET
// Description: This example shows how to build a workbook, add sample date/value rows, create a pivot table, attach a printable timeline to the Date field, adjust its shape, and save the workbook as a PDF where the timeline repeats on each generated page.
// Keywords: Aspose.Cells | C# timeline | pivot table timeline | PDF export Aspose.Cells | printable timeline | .NET workbook to PDF | repeat timeline each page | PdfSaveOptions timeline
// Common Searches: Aspose.Cells add timeline to pivot table | timeline printable on each PDF page Aspose.Cells | export workbook with timeline to PDF .NET | repeat timeline on every page PDF Aspose | how to set timeline shape printable Aspose.Cells
// Developer Intent: Add a timeline linked to a pivot table and ensure it appears on every page of the PDF generated with Aspose.Cells for .NET.
// Use Cases: Sales report PDF with a date timeline header on each page for consistent context. | Multi‑page financial workbook where a quarterly timeline repeats on every page. | Project schedule PDF that shows an overall timeline banner on all pages.
// AI Prompts: Generate C# code using Aspose.Cells to create a pivot table, add a printable timeline, and export the workbook to PDF with the timeline on each page. | Explain how PdfSaveOptions can be configured so a timeline repeats on every PDF page in Aspose.Cells. | Show how to set the size, position, and caption of a timeline shape before saving the workbook as a PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace TimelinePdfDemo
{
    // This example shows how to build a workbook, add sample date/value rows, create a pivot table, attach a printable timeline to the Date field, adjust its shape, and save the workbook as a PDF where the timeline repeats on each generated page.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate worksheet with sample data (date and value)
                cells["A1"].Value = "Date";
                cells["B1"].Value = "Value";

                cells["A2"].Value = new DateTime(2023, 1, 1);
                cells["B2"].Value = 1200;

                cells["A3"].Value = new DateTime(2023, 2, 1);
                cells["B3"].Value = 1500;

                cells["A4"].Value = new DateTime(2023, 3, 1);
                cells["B4"].Value = 1800;

                // Create a pivot table that will serve as the data source for the timeline
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                // Add the date field to the Page area (required for timeline)
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                // Add the value field to the Data area
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table using the date field as the base field
                int timelineIdx = sheet.Timelines.Add(pivot, "A20", "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Adjust timeline shape size and position
                timeline.Shape.Top = 100;      // pixels from top of the sheet
                timeline.Shape.Left = 10;      // pixels from left of the sheet
                timeline.Shape.Width = 500;    // width in pixels
                timeline.Shape.Height = 80;    // height in pixels
                timeline.Shape.IsPrintable = true; // ensure it is printed on each page

                // Optional: set a caption for visual identification
                timeline.Caption = "Sales Timeline";

                // Configure PDF save options (default settings will keep the timeline on each printed page)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = false // allow multiple pages; timeline repeats on each page
                };

                // Save the workbook as PDF
                workbook.Save("TimelineOnEachPage.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
