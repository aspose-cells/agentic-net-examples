using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // Add a hidden worksheet
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Hide a column and a row in the visible sheet (optional, also excluded)
            visibleSheet.Cells.HideColumn(1); // Hide column B
            visibleSheet.Cells.HideRow(2);    // Hide row 3 (zero‑based index)

            // Configure HTML save options to exclude hidden elements
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false,                     // Do not export hidden worksheets
                HiddenColDisplayType = HtmlHiddenColDisplayType.Remove, // Remove hidden columns from HTML
                HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove  // Remove hidden rows from HTML
            };

            // Save the workbook as HTML with the specified options
            workbook.Save("output_without_hidden_elements.html", saveOptions);
        }
    }
}