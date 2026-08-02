// Title: Export Merged Worksheet Titles to a Single HTML File with Aspose.Cells for .NET
// Description: Demonstrates how to insert each worksheet's name as a merged, styled heading row, then save all sheets into one HTML document using HtmlSaveOptions (ExportRowColumnHeadings, SaveAsSingleFile, ShowAllSheets).
// Keywords: Aspose.Cells C# HTML export merged heading | ExportRowColumnHeadings .NET | HtmlSaveOptions ShowAllSheets | single HTML file multiple worksheets | worksheet name as heading Aspose | C# Aspose.Cells HTMLSaveOptions example | merge cells for title Excel to HTML | Aspose.Cells HTML report generation | export Excel workbook to single HTML
// Common Searches: Aspose.Cells export worksheet name as heading HTML C# | HtmlSaveOptions ExportRowColumnHeadings example | Save all Excel sheets to one HTML file Aspose | How to merge cells for sheet title in HTML export | Create single HTML report from multiple worksheets Aspose.Cells
// Developer Intent: Generate one HTML file where every worksheet appears with its name displayed as a merged, bold heading row at the top of its table.
// Use Cases: Produce a printable web report that consolidates several worksheets, each clearly labeled with a section heading. | Create an email‑ready HTML document that preserves the original sheet titles for quick reference. | Build a web‑viewable spreadsheet where users can scroll through all sheets without losing context of each sheet's purpose.
// AI Prompts: Add a background color to the merged worksheet heading while keeping ExportRowColumnHeadings enabled. | Show a version that uses only ExportRowColumnHeadings (no obsolete ExportHeadings) to achieve the same output. | Explain how to export each worksheet to separate HTML files but still insert the sheet name as a merged heading row. | Provide code to customize the heading font (color, size) for all sheets in the HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to insert each worksheet's name as a merged, styled heading row, then save all sheets into one HTML document using HtmlSaveOptions (ExportRowColumnHeadings, SaveAsSingleFile, ShowAllSheets).
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sales");
            workbook.Worksheets.Add("Inventory");

            // Populate each worksheet with sample data
            PopulateWorksheet(workbook.Worksheets[0]); // Sheet1 (default name "Sheet1")
            PopulateWorksheet(workbook.Worksheets[1]); // "Sales"
            PopulateWorksheet(workbook.Worksheets[2]); // "Inventory"

            // Insert the worksheet name as a heading in the first row of each sheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Shift existing data down by one row
                sheet.Cells.InsertRows(0, 1);
                // Write the sheet name as a heading (merged across columns A‑C)
                sheet.Cells["A1"].PutValue(sheet.Name);
                sheet.Cells.Merge(0, 0, 1, 3);
                // Apply a simple style to make the heading stand out
                Style headingStyle = workbook.CreateStyle();
                headingStyle.Font.IsBold = true;
                headingStyle.Font.Size = 14;
                headingStyle.HorizontalAlignment = TextAlignmentType.Center;
                sheet.Cells["A1"].SetStyle(headingStyle);
            }

            // Configure HTML save options to export row/column headings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // ExportHeadings is obsolete but still works; the recommended property is ExportRowColumnHeadings
            htmlOptions.ExportHeadings = true;
            htmlOptions.ExportRowColumnHeadings = true;

            // Save the workbook as a single HTML file that contains all sheets
            htmlOptions.SaveAsSingleFile = true;
            htmlOptions.ShowAllSheets = true; // ensure all sheets are included in the single file

            workbook.Save("WorkbookWithSheetHeadings.html", htmlOptions);
        }

        // Helper method to add some dummy data to a worksheet
        private static void PopulateWorksheet(Worksheet sheet)
        {
            sheet.Cells["A2"].PutValue("Header1");
            sheet.Cells["B2"].PutValue("Header2");
            sheet.Cells["C2"].PutValue("Header3");

            sheet.Cells["A3"].PutValue("Data1");
            sheet.Cells["B3"].PutValue("Data2");
            sheet.Cells["C3"].PutValue("Data3");
        }
    }
}
