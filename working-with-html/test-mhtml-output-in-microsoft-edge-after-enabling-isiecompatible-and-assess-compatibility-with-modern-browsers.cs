using System;
using Aspose.Cells;

namespace AsposeCellsMHtmlTest
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("MHTML Compatibility Test");
            sheet.Cells["B2"].PutValue(DateTime.Now);
            sheet.Cells["C3"].PutValue(12345);

            // Configure HTML save options for MHTML with IE compatibility enabled
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.IsIECompatible = true; // Enable IE compatibility mode

            // Save the workbook as MHTML (MHtml) format
            workbook.Save("MHtmlOutput.mht", saveOptions);
            Console.WriteLine("MHTML file saved with IsIECompatible = true.");

            // Load the saved MHTML file to verify it can be read back
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.MHtml);
            Workbook loadedWorkbook = new Workbook("MHtmlOutput.mht", loadOptions);

            // Output a cell value from the loaded workbook to confirm successful load
            string loadedValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Loaded cell A1 value: " + loadedValue);
        }
    }
}