using System;
using Aspose.Cells;

class ExportExcelToHtmlWithNavigation
{
    static void Main()
    {
        // Create a new workbook and add some worksheets
        Workbook workbook = new Workbook();

        // First worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Summary";
        sheet1.Cells["A1"].PutValue("Summary Sheet Content");

        // Second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("Details");
        sheet2.Cells["A1"].PutValue("Details Sheet Content");

        // Third worksheet
        Worksheet sheet3 = workbook.Worksheets.Add("Report");
        sheet3.Cells["A1"].PutValue("Report Sheet Content");

        // Configure HTML save options to generate a navigation pane (tabstrip)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export the whole workbook (default is false, set explicitly for clarity)
        saveOptions.ExportActiveWorksheetOnly = false;

        // Keep SaveAsSingleFile = false (default) so Aspose.Cells creates separate HTML files
        // and a tabstrip.htm file that acts as the navigation pane.
        // Optionally include hidden worksheets in the navigation
        saveOptions.ExportHiddenWorksheet = true;

        // Save the workbook to HTML. The output will contain:
        // - tabstrip.htm (navigation pane)
        // - sheet0.htm, sheet1.htm, sheet2.htm (individual sheet pages)
        // - filelist.xml (list of generated files)
        workbook.Save("WorkbookWithNavigation.html", saveOptions);
    }
}