// Title: Export workbook to HTML with Aspose.Cells, hide rows, columns and tables, and verify only visible TableCssId prefixes appear
// AI Prompts: Generate C# code that creates two ListObjects in a worksheet, hides the rows and columns of the second table, and saves the workbook to HTML using Aspose.Cells. | Show how to configure HtmlSaveOptions so that hidden worksheets, rows, columns, and tables are omitted from the HTML output. | Write a C# verification step that checks the generated HTML contains the display name of the visible table but not the hidden table's CSS identifier.
// Common Searches: Aspose.Cells export to HTML without hidden rows and columns C# | How to hide a table in Excel and exclude it from HTML output using Aspose.Cells | Verify TableCssId appears only for visible tables in Aspose.Cells HTML export | C# HtmlSaveOptions ExportHiddenWorksheet false example with hidden tables | Exclude hidden ListObject from HTML when saving workbook with Aspose.Cells
// Tags: HTML export excluding hidden rows Aspose.Cells | hide columns before saving workbook to HTML C# | visible ListObject TableCssId verification Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet false usage | validate hidden table not rendered in HTML Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook with two ListObjects, hides the rows and columns containing the second table, saves the workbook to HTML using HtmlSaveOptions with ExportHiddenWorksheet set to false, and then checks that the generated HTML includes only the visible table's display name, confirming hidden tables are omitted.
class ExportHiddenRowsColumns
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate data for two tables
            // Table 1 (visible)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Table 2 (will be hidden)
            sheet.Cells["D1"].PutValue("Code");
            sheet.Cells["E1"].PutValue("Value");
            sheet.Cells["D2"].PutValue("X");
            sheet.Cells["E2"].PutValue(100);
            sheet.Cells["D3"].PutValue("Y");
            sheet.Cells["E3"].PutValue(200);

            // Create ListObject (table) for visible data
            int firstRowVisible = 0; // zero‑based index
            int firstColumnVisible = 0;
            int totalRowsVisible = 3;
            int totalColumnsVisible = 2;
            ListObject visibleTable = sheet.ListObjects[sheet.ListObjects.Add(
                firstRowVisible,
                firstColumnVisible,
                firstRowVisible + totalRowsVisible,
                firstColumnVisible + totalColumnsVisible,
                true)];
            // Set a display name that will appear in the HTML output
            visibleTable.DisplayName = "tblVisible";
            visibleTable.TableStyleType = TableStyleType.TableStyleMedium2;

            // Create ListObject (table) for hidden data
            int firstRowHidden = 0;
            int firstColumnHidden = 3; // column D
            int totalRowsHidden = 3;
            int totalColumnsHidden = 2;
            ListObject hiddenTable = sheet.ListObjects[sheet.ListObjects.Add(
                firstRowHidden,
                firstColumnHidden,
                firstRowHidden + totalRowsHidden,
                firstColumnHidden + totalColumnsHidden,
                true)];
            hiddenTable.DisplayName = "tblHidden";
            hiddenTable.TableStyleType = TableStyleType.TableStyleMedium9;

            // Hide rows 2‑3 (index 1‑2) and columns D and E (indexes 3‑4) that contain the hidden table
            sheet.Cells.HideRow(1);
            sheet.Cells.HideRow(2);
            sheet.Cells.HideColumn(3);
            sheet.Cells.HideColumn(4);

            // Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportHiddenWorksheet = false // Do not export hidden worksheets
            };

            // Save workbook to a memory stream as HTML
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0;
                string htmlContent = new StreamReader(htmlStream, Encoding.UTF8).ReadToEnd();

                // Verify that only the visible table's display name appears
                bool visiblePresent = htmlContent.Contains("tblVisible");
                bool hiddenPresent = htmlContent.Contains("tblHidden");

                Console.WriteLine("Visible table display name present: " + visiblePresent);
                Console.WriteLine("Hidden table display name present: " + hiddenPresent);
                // Expected output: true for visible, false for hidden
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
