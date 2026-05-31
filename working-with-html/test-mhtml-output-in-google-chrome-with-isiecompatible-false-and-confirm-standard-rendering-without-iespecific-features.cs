using System;
using Aspose.Cells;

class TestMhtmlOutput
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to the worksheet
        worksheet.Cells["A1"].PutValue("MHTML Output Test");
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Disable IE‑specific compatibility mode
        saveOptions.IsIECompatible = false;

        // Use HTML5 standard for modern browsers like Google Chrome
        saveOptions.HtmlVersion = HtmlVersion.Html5;

        // Save the workbook as MHTML (MHT) using the configured options
        workbook.Save("MhtmlOutput.mht", saveOptions);

        Console.WriteLine("MHTML file saved with IsIECompatible = false (standard rendering).");
    }
}