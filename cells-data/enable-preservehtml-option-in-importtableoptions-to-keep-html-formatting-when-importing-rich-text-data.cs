// Title: Import a DataTable containing HTML links into Excel while preserving HTML formatting with Aspose.Cells for .NET
// AI Prompts: Generate C# code that imports a DataTable with HTML strings into an Excel worksheet using Aspose.Cells and retains the HTML tags. | Show how to enable ImportTableOptions.IsHtmlString and apply custom number formats when importing rich‑text data with Aspose.Cells. | Demonstrate saving the workbook after importing HTML content to an .xlsx file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# import DataTable preserve HTML tags | How to keep HTML formatting when using Cells.ImportData in .NET | ImportTableOptions IsHtmlString example for Excel export | Saving HTML links from a DataTable to Excel with Aspose.Cells | Set time column number format while importing HTML content using Aspose.Cells
// Tags: importdatatable preservehtml aspocells | importtableoptions ishtmlstring property | excel export html content c# | cells.importdata html formatting | numberformats time column aspocells

using System;
using System.Data;
using Aspose.Cells;

// The example creates a DataTable with an HTML anchor tag, configures ImportTableOptions to set IsHtmlString=true (and optional time number formats), imports the table into the first worksheet starting at cell A1, and saves the workbook as PreserveHtmlImportDemo.xlsx, preserving the HTML formatting in the Excel cells.
class PreserveHtmlImportDemo
{
    static void Main()
    {
        // Create a DataTable with HTML content
        DataTable table = new DataTable();
        table.Columns.Add("ID");
        table.Columns.Add("Content");
        table.Columns.Add("Time");

        // Add a row containing an HTML link
        table.Rows.Add("1", "<a href='https://www.example.com'>Example Link</a>", "2:30 PM");

        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure import options to preserve HTML formatting
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true,               // import column headers
            IsHtmlString = true,                   // preserve HTML tags
            NumberFormats = new string[] { null, null, "h:mm AM/PM" } // optional time format
        };

        // Import the DataTable starting at cell A1 (row 0, column 0)
        cells.ImportData(table, 0, 0, importOptions);

        // Save the workbook to an Excel file
        workbook.Save("PreserveHtmlImportDemo.xlsx");
    }
}
