// Title: Create a weekly timeline with bold headers in an Excel workbook and export it as a PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that fills a worksheet with weekly dates, applies a bold style to the header row, builds a pivot table, adds a timeline control linked to the Date field, and saves the workbook to PDF with Arial font embedding and OnePagePerSheet rendering. | Adjust the Aspose.Cells example to use a custom date format (mm/dd/yyyy) and configure PdfSaveOptions so the generated PDF embeds standard Windows fonts and places each worksheet on a single page. | Enhance the sample by adding a second timeline for an additional date column and produce a multi‑page PDF where each page displays a different timeline view.
// Common Searches: how to generate a weekly timeline in Excel using Aspose.Cells C# | Aspose.Cells add timeline control to pivot table and export to PDF | C# set bold font for header row in Aspose.Cells workbook | export Excel sheet with timeline to PDF with OnePagePerSheet option | custom date format mm/dd/yyyy in Aspose.Cells worksheet
// Tags: Aspose.Cells timeline control creation | Aspose.Cells bold header style C# | Aspose.Cells pivot table weekly dates | Aspose.Cells PDF export OnePagePerSheet | Aspose.Cells embed fonts PDF rendering

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelinePdfDemo
{
    // Demonstrates populating an Excel sheet with weekly dates, applying bold headers, creating a pivot table, linking a timeline control, and exporting the workbook to a single‑page‑per‑sheet PDF with embedded fonts using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -----------------------------------------------------------------
            // 1. Populate worksheet with weekly dates and sample values
            // -----------------------------------------------------------------
            // Header row (bold)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;

            cells["A1"].PutValue("Date");
            cells["A1"].SetStyle(headerStyle);
            cells["B1"].PutValue("Value");
            cells["B1"].SetStyle(headerStyle);

            // Date style (mm/dd/yyyy)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm/dd/yyyy";

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++) // 10 weeks of data
            {
                // Date column
                cells[i + 1, 0].PutValue(startDate.AddDays(i * 7));
                cells[i + 1, 0].SetStyle(dateStyle);

                // Value column (sample numeric data)
                cells[i + 1, 1].PutValue(i * 10 + 5);
            }

            // -----------------------------------------------------------------
            // 2. Create a PivotTable based on the data
            // -----------------------------------------------------------------
            // Add the pivot table starting at cell D1
            int pivotIndex = sheet.PivotTables.Add("A1:B11", "D1", "WeeklyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Row field: Date
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            // Data field: Value (sum)
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh to calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // -----------------------------------------------------------------
            // 3. Add a Timeline control linked to the PivotTable
            // -----------------------------------------------------------------
            // Place the timeline at row 0, column 5 (cell F1) and bind it to the "Date" field
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Set a caption for the timeline (the caption inherits the worksheet's default font,
            // which is already bold for the header row; additional styling of the caption
            // itself is not directly exposed, so we rely on the default appearance.)
            timeline.Caption = "Weekly Timeline";

            // Optional visual settings
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;

            // -----------------------------------------------------------------
            // 4. Save the workbook as a PDF document
            // -----------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each worksheet is rendered on a single page
                OnePagePerSheet = true,
                // Use a default font that supports the date format
                DefaultFont = "Arial",
                // Embed standard Windows fonts for better compatibility
                EmbedStandardWindowsFonts = true
            };

            // Save the workbook (which includes the timeline) to PDF
            workbook.Save("WeeklyTimeline.pdf", pdfOptions);
        }
    }
}
