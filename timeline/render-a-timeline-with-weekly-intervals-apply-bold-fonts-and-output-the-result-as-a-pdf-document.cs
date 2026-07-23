// Title: Export a Weekly PivotTable Timeline to PDF with Bold Caption using Aspose.Cells for .NET
// Description: This example creates a workbook, fills column A with eight consecutive weekly dates and column B with numeric values, applies bold styling to the header, formats the date column, builds a PivotTable, adds a Timeline control linked to the "Week" field with a bold caption, and saves the sheet as a PDF with one page per sheet and embedded fonts.
// Keywords: Aspose.Cells | C# timeline PDF | PivotTable timeline | bold caption Aspose.Cells | date format Excel C# | PdfSaveOptions Aspose | export worksheet to PDF | .NET Excel timeline
// Common Searches: Aspose.Cells add timeline to pivot table C# | Render timeline and export to PDF Aspose.Cells | Bold caption for timeline control .NET | Save Excel workbook as PDF with timeline | Weekly timeline example Aspose.Cells
// Developer Intent: Generate a weekly timeline bound to a PivotTable, apply bold formatting to its caption, and export the worksheet as a PDF document using Aspose.Cells for .NET.
// Use Cases: Weekly sales dashboard: display sales data in a PivotTable, filter weeks via a timeline, and share the view as a PDF report. | Project schedule: list weekly milestones, let stakeholders select periods with a timeline, and produce a PDF summary for review.
// AI Prompts: Write C# code that adds a Timeline to an existing PivotTable and sets the caption font to bold with Aspose.Cells. | Show how to configure PdfSaveOptions so a timeline and its PivotTable appear on the same PDF page. | Explain how to apply a custom date format to a column of dates in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelinePdfDemo
{
    // This example creates a workbook, fills column A with eight consecutive weekly dates and column B with numeric values, applies bold styling to the header, formats the date column, builds a PivotTable, adds a Timeline control linked to the "Week" field with a bold caption, and saves the sheet as a PDF with one page per sheet and embedded fonts.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate sample data with weekly dates
            // ------------------------------------------------------------
            // Header row
            cells["A1"].PutValue("Week");
            cells["B1"].PutValue("Value");

            // Apply bold font to the header
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            cells["A1"].SetStyle(headerStyle);
            cells["B1"].SetStyle(headerStyle);

            // Generate 8 weeks of data starting from today
            DateTime startDate = DateTime.Today;
            for (int i = 0; i < 8; i++)
            {
                // Column A: week start date
                cells[i + 1, 0].PutValue(startDate.AddDays(i * 7));
                // Column B: some numeric value
                cells[i + 1, 1].PutValue(10 + i * 5);
            }

            // Apply a date style to the date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";
            for (int i = 0; i < 8; i++)
            {
                cells[i + 1, 0].SetStyle(dateStyle);
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable based on the data
            // ------------------------------------------------------------
            // Define the source range (including header)
            string sourceRange = "A1:B9";
            // Destination cell for the pivot table
            string destCell = "D1";

            int pivotIndex = sheet.PivotTables.Add(sourceRange, destCell, "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the date field to the row area and the value field to the data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Week");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh the pivot table so that it contains data
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a Timeline control linked to the PivotTable
            // ------------------------------------------------------------
            // The timeline will be placed starting at row 12, column 0 (cell A12)
            int timelineIndex = sheet.Timelines.Add(pivot, 11, 0, "Week");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Set a caption for the timeline
            timeline.Caption = "Weekly Timeline";

            // Apply bold formatting to the timeline caption.
            // The TimelineShape exposes a Font property that can be used to style the caption.
            // (If the API version does not expose Font directly, this line can be omitted.)
            timeline.Shape.Font.IsBold = true;

            // ------------------------------------------------------------
            // 4. Save the workbook as a PDF document
            // ------------------------------------------------------------
            // Configure PDF save options (optional)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the timeline is rendered on the same page as the pivot table
                OnePagePerSheet = true,
                // Use a default font that supports the date format
                DefaultFont = "Arial",
                // Embed standard Windows fonts
                EmbedStandardWindowsFonts = true
            };

            // Save to PDF
            string pdfPath = "WeeklyTimeline.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Timeline rendered and saved to PDF: {pdfPath}");
        }
    }
}
