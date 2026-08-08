// Title: Add a Clickable Hyperlink to an Aspose.Cells Timeline Shape that Opens a Different Worksheet in a PDF (C#)
// Description: This example builds a workbook with a data sheet, a pivot table, and a timeline linked to the pivot. It then attaches a hyperlink to the timeline shape so that clicking the timeline in the exported PDF jumps to cell A1 of a second worksheet named "Details". The hyperlink’s address, display text, and screen tip are configured via the shape’s Hyperlink property, and the PDF retains the active link.
// Keywords: Aspose.Cells timeline hyperlink C# | timeline shape link to worksheet | export PDF with active links Aspose.Cells | clickable timeline Aspose.Cells .NET | hyperlink timeline shape PDF | Aspose.Cells PDF navigation | C# Aspose.Cells timeline example
// Common Searches: how to add hyperlink to timeline shape Aspose.Cells | Aspose.Cells timeline click opens another sheet PDF | C# export workbook to PDF with active timeline link | Aspose.Cells timeline navigation between worksheets | add clickable timeline marker in Aspose.Cells PDF
// Developer Intent: Create a PDF where clicking a timeline element navigates directly to a specified worksheet within the same workbook.
// Use Cases: Interactive sales dashboard PDF where the timeline links to a detailed data sheet for each period. | Project timeline report that opens the corresponding milestone worksheet when a user clicks a timeline marker. | Financial summary PDF that lets readers jump from a high‑level chart to supporting schedules via timeline hyperlinks.
// AI Prompts: Generate C# code using Aspose.Cells to attach a hyperlink to a TimelineShape that opens a target worksheet when the PDF is opened. | Explain how to set Hyperlink.Address, TextToDisplay, and ScreenTip on a TimelineShape so the link works after PDF conversion. | Show steps to verify that a timeline hyperlink remains functional in the exported PDF file.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelineHyperlinkDemo
{
    // This example builds a workbook with a data sheet, a pivot table, and a timeline linked to the pivot. It then attaches a hyperlink to the timeline shape so that clicking the timeline in the exported PDF jumps to cell A1 of a second worksheet named "Details". The hyperlink’s address, display text, and screen tip are configured via the shape’s Hyperlink property, and the PDF retains the active link.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet setup
                // -------------------------------------------------
                // First worksheet will contain the data, pivot table and timeline
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Second worksheet will be the target of the hyperlink
                Worksheet detailSheet = workbook.Worksheets.Add("Details");
                detailSheet.Cells["A1"].PutValue("Details Sheet - Clicked from Timeline");

                // -------------------------------------------------
                // Populate sample data for the pivot table
                // -------------------------------------------------
                Cells cells = dataSheet.Cells;
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Date");
                cells["C1"].PutValue("Sales");

                cells["A2"].PutValue("P1");
                cells["B2"].PutValue(new DateTime(2023, 1, 1));
                cells["C2"].PutValue(100);

                cells["A3"].PutValue("P2");
                cells["B3"].PutValue(new DateTime(2023, 1, 2));
                cells["C3"].PutValue(150);

                cells["A4"].PutValue("P1");
                cells["B4"].PutValue(new DateTime(2023, 2, 1));
                cells["C4"].PutValue(200);

                // -------------------------------------------------
                // Create a pivot table based on the data
                // -------------------------------------------------
                PivotTableCollection pivots = dataSheet.PivotTables;
                int pivotIndex = pivots.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivot = pivots[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Column, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // Add a timeline linked to the pivot table
                // -------------------------------------------------
                // Place the timeline at cell A10
                int timelineIndex = dataSheet.Timelines.Add(pivot, "A10", "Date");
                Timeline timeline = dataSheet.Timelines[timelineIndex];

                // -------------------------------------------------
                // Add an interactive hyperlink to the timeline shape
                // -------------------------------------------------
                TimelineShape timelineShape = (TimelineShape)timeline.Shape;

                // Add a hyperlink to the worksheet; the shape will reference it automatically
                int linkIdx = dataSheet.Hyperlinks.Add("A10", 1, 1, "Details!A1");
                Hyperlink hyperlink = dataSheet.Hyperlinks[linkIdx];

                // Configure hyperlink properties via the shape's Hyperlink reference
                if (timelineShape.Hyperlink != null)
                {
                    timelineShape.Hyperlink.Address = hyperlink.Address;
                    timelineShape.Hyperlink.TextToDisplay = "Open Details Sheet";
                    timelineShape.Hyperlink.ScreenTip = "Click to navigate to Details sheet";
                }

                // -------------------------------------------------
                // Save the workbook as PDF – the hyperlink will be active in the PDF
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save("TimelineWithHyperlink.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
