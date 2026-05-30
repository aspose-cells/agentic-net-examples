using System;
using Aspose.Cells;

class GenerateMhtmlIECompatible
{
    static void Main()
    {
        // Create a new workbook with multiple worksheets
        Workbook workbook = new Workbook();

        // First sheet (default)
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[0].Cells["A1"].PutValue("Content of Sheet 1");

        // Second sheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Content of Sheet 2");

        // Third sheet
        Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
        sheet3.Cells["A1"].PutValue("Content of Sheet 3");

        // Configure HTML save options for IE compatibility and tab switching
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.IsIECompatible = true;      // Enable IE‑compatible output
        saveOptions.ShowAllSheets = true;       // Export all sheets as tabs
        saveOptions.SaveAsSingleFile = true;    // Produce a single MHTML file

        // Save the workbook as MHTML (MHT) using the configured options
        workbook.Save("MultiSheet_IE_Compatible.mht", saveOptions);

        // Inform the user; opening the resulting file in Internet Explorer should show tab navigation
        Console.WriteLine("MHTML file saved. Open it in Internet Explorer to verify tab switching.");
    }
}