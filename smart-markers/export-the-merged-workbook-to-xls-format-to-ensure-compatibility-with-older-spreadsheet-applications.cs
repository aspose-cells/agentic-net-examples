// Title: Export a merged Aspose.Cells workbook to XLS (Excel 97‑2003) with C#
// Description: Demonstrates how to create or load a merged workbook, configure XlsSaveOptions, and save it as an Excel 97‑2003 (XLS) file for legacy spreadsheet compatibility using Aspose.Cells for .NET.
// Keywords: Aspose.Cells export XLS | C# XlsSaveOptions | Excel 97-2003 compatibility | merged workbook to XLS | legacy Excel export | Aspose.Cells .NET example
// Common Searches: save Aspose.Cells workbook as XLS in C# | export merged workbook to Excel 97-2003 format | XlsSaveOptions example Aspose.Cells | convert .xlsx to .xls using Aspose.Cells | C# code for backward‑compatible Excel export
// Developer Intent: Generate an XLS file from a merged workbook to support older Excel versions.
// Use Cases: Produce a single report that combines data from several worksheets and deliver it to users still on Excel 2003. | Automate archival of merged workbooks in XLS format for long‑term storage in environments without newer Office suites. | Distribute programmatically created financial statements to clients who require the legacy Excel file type.
// AI Prompts: Write C# code that loads an existing merged workbook, applies XlsSaveOptions, and saves it as an XLS file while handling I/O exceptions. | Explain step‑by‑step how to configure XlsSaveOptions for maximum compatibility when exporting a workbook that contains smart markers. | Create a sample that merges multiple worksheets into one workbook and then exports the result to XLS using Aspose.Cells in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExportToXls
{
    // Demonstrates how to create or load a merged workbook, configure XlsSaveOptions, and save it as an Excel 97‑2003 (XLS) file for legacy spreadsheet compatibility using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create or load the merged workbook.
            // For demonstration, we create a new workbook and add some data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Initialize XlsSaveOptions for Excel 97-2003 format.
            XlsSaveOptions saveOptions = new XlsSaveOptions();

            // Save the workbook as an XLS file using the save options.
            workbook.Save("MergedWorkbook.xls", saveOptions);

            Console.WriteLine("Workbook exported successfully to XLS format.");
        }
    }
}
