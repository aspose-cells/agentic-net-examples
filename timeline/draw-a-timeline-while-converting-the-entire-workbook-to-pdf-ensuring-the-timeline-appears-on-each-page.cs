// Title: Add a Pivot Timeline and Export Workbook to PDF with Timeline on Every Page – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate it with date‑based data, build a pivot table, attach a timeline control, adjust its shape, and save the file as a PDF where the timeline appears on each page that contains its area.
// Keywords: Aspose.Cells timeline PDF | C# pivot table timeline | export workbook to PDF Aspose.Cells | timeline shape properties | PDFSaveOptions timeline rendering | multi‑page PDF timeline Aspose | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add timeline to pivot table | export timeline to PDF with Aspose.Cells | timeline appears on every PDF page | set timeline size and position C# | Aspose.Cells PDFSaveOptions timeline
// Developer Intent: Create a PDF from a workbook that includes a pivot‑driven timeline visible on each page.
// Use Cases: Sales report PDF with a persistent timeline filter for quick date navigation. | Project schedule PDF where a timeline linked to a pivot table lets readers scroll through months across pages. | Financial statement PDF that retains a timeline control for selecting fiscal periods on every page.
// AI Prompts: Generate C# code using Aspose.Cells to add a timeline linked to a pivot table and ensure it renders on every page of the exported PDF. | Explain how to modify the timeline's shape dimensions and placement before saving a workbook as PDF with Aspose.Cells. | Provide troubleshooting steps when a timeline does not appear on all PDF pages after export with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

namespace TimelinePdfDemo
{
    // Demonstrates how to create a workbook, populate it with date‑based data, build a pivot table, attach a timeline control, adjust its shape, and save the file as a PDF where the timeline appears on each page that contains its area.
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

                // Populate sample data with a date field (required for a timeline)
                cells["A1"].Value = "Date";
                cells["B1"].Value = "Sales";

                cells["A2"].Value = new DateTime(2023, 1, 1);
                cells["B2"].Value = 1200;

                cells["A3"].Value = new DateTime(2023, 2, 1);
                cells["B3"].Value = 1500;

                cells["A4"].Value = new DateTime(2023, 3, 1);
                cells["B4"].Value = 1800;

                // Create a pivot table that will serve as the data source for the timeline
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table (using cell address for placement)
                int timelineIdx = sheet.Timelines.Add(pivot, "A20", "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Adjust the visual appearance of the timeline via its Shape object
                timeline.Shape.Top = 200;      // vertical offset in pixels
                timeline.Shape.Left = 10;      // horizontal offset in pixels
                timeline.Shape.Width = 600;    // width in pixels
                timeline.Shape.Height = 80;    // height in pixels
                timeline.Caption = "Sales Timeline";

                // Prepare PDF save options – keep default pagination so the timeline appears on each page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    CreatedTime = DateTime.Now
                };

                // Save the workbook as PDF; the timeline will be rendered on every page where its area lies
                workbook.Save("TimelineReport.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
