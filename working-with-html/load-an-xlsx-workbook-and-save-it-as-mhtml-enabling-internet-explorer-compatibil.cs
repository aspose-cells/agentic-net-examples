using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsMhtmlExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook from disk
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options for MHTML and enable Internet Explorer compatibility
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
            saveOptions.IsIECompatible = true; // Enable IE compatibility mode

            // Save the workbook as an MHTML file using the configured options
            workbook.Save("output.mht", saveOptions);

            Console.WriteLine("Workbook successfully saved as MHTML with IE compatibility.");
        }
    }
}