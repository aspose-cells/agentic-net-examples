// Title: C# – Create a Weekly Sales Timeline with Bold Headers and Export to PDF using Aspose.Cells
// Description: This example shows how to build a new workbook, add a bold header row, populate eight weeks of sales data, create a pivot table, attach a timeline control to the Date field, make the timeline caption bold, and save the sheet as a single‑page PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline PDF | C# Aspose.Cells pivot table | weekly date series Aspose.Cells | bold header cells Aspose.Cells | export worksheet to PDF Aspose | timeline control Aspose.Cells | Aspose.Cells PDFSaveOptions
// Common Searches: how to add a timeline to a pivot table in Aspose.Cells C# | export worksheet with timeline to PDF using Aspose.Cells | set bold font for header cells Aspose.Cells | create weekly date series in Excel with Aspose.Cells | Aspose.Cells PDFSaveOptions one page per sheet
// Developer Intent: Generate a workbook with weekly sales data, apply bold formatting to headers and timeline caption, link a timeline to a pivot table, and save the result as a PDF.
// Use Cases: Produce a PDF report that visualizes weekly sales trends with a filterable timeline for executive reviews. | Automate weekly sales dashboards where the header row and timeline caption are emphasized in bold for printed distribution. | Integrate timeline‑driven PDF generation into a .NET backend for scheduled email delivery of sales performance summaries.
// AI Prompts: Show how to change the timeline caption font size and color in Aspose.Cells C#. | Provide code to add multiple timeline fields to a pivot table and export the workbook to PDF. | Explain how to customize PDF page margins and orientation when saving a workbook that contains a timeline.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace AsposeCellsTimelinePdfDemo
{
    // This example shows how to build a new workbook, add a bold header row, populate eight weeks of sales data, create a pivot table, attach a timeline control to the Date field, make the timeline caption bold, and save the sheet as a single‑page PDF with Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // Populate sample data with weekly dates
                // -------------------------------------------------
                // Header row
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                // Apply bold font to header cells
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                StyleFlag headerFlag = new StyleFlag { FontBold = true };
                cells["A1"].SetStyle(headerStyle, headerFlag);
                cells["B1"].SetStyle(headerStyle, headerFlag);

                // Starting date
                DateTime startDate = new DateTime(2023, 1, 1);
                // Fill 8 weeks of data
                for (int i = 0; i < 8; i++)
                {
                    // Date column (weekly interval)
                    cells[i + 1, 0].PutValue(startDate.AddDays(i * 7));
                    // Sample sales value
                    cells[i + 1, 1].PutValue(1000 + i * 250);
                }

                // -------------------------------------------------
                // Create a PivotTable based on the data range
                // -------------------------------------------------
                // Data range: A1:B9 (header + 8 rows)
                string dataRange = "A1:B9";
                // Destination cell for the pivot table
                string pivotDest = "D1";
                int pivotIndex = sheet.PivotTables.Add(dataRange, pivotDest, "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                // Date field must be added as a Page (filter) field for Timeline support
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // Add a Timeline control linked to the PivotTable
                // -------------------------------------------------
                // Place the timeline at row 12, column 0 (cell A12)
                int timelineIndex = sheet.Timelines.Add(pivot, 11, 0, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // Set timeline caption
                timeline.Caption = "Weekly Sales Timeline";

                // Attempt to make the caption bold (if supported)
                try
                {
                    var shape = timeline.Shape;
                    var fontProp = shape.GetType().GetProperty("Font");
                    if (fontProp != null)
                    {
                        var font = fontProp.GetValue(shape, null);
                        var boldProp = font?.GetType().GetProperty("Bold");
                        boldProp?.SetValue(font, true, null);
                    }
                }
                catch
                {
                    // Ignore if the shape does not expose a Font property
                }

                // -------------------------------------------------
                // Save the workbook as a PDF document
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };
                workbook.Save("WeeklyTimeline.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
