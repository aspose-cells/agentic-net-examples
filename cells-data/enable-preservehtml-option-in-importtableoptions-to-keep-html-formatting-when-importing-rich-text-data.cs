// Title: Preserve HTML Formatting When Importing a DataTable into Excel with Aspose.Cells .NET
// Description: Demonstrates how to set the IsHtmlString flag in ImportTableOptions so that HTML tags (e.g., hyperlinks, bold, italic) are kept as formatted cell content when a DataTable is imported into an Aspose.Cells worksheet and saved as .xlsx.
// Keywords: Aspose.Cells ImportTableOptions | IsHtmlString | HTML import Excel .NET | preserve HTML in Excel | ImportData HTML formatting | C# Aspose.Cells rich text
// Common Searches: Aspose.Cells preserve HTML when importing | ImportTableOptions IsHtmlString example | C# import DataTable with HTML tags into Excel | keep hyperlinks after ImportData Aspose.Cells | how to retain HTML styling in Excel cells
// Developer Intent: Enable the IsHtmlString flag so HTML markup is rendered as formatted content rather than plain text during DataTable import.
// Use Cases: Import a report that contains clickable links and styled text directly from a database. | Migrate web‑generated HTML summaries into Excel while preserving visual formatting. | Create dashboards where comments or descriptions include bold, italic, or colored HTML snippets.
// AI Prompts: Show C# code that imports a DataTable with HTML strings into an Aspose.Cells worksheet using ImportTableOptions.IsHtmlString. | Explain how to configure ImportTableOptions to retain HTML formatting and save the workbook. | What steps are needed to ensure HTML entities render correctly after importing with Aspose.Cells?

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    // Demonstrates how to set the IsHtmlString flag in ImportTableOptions so that HTML tags (e.g., hyperlinks, bold, italic) are kept as formatted cell content when a DataTable is imported into an Aspose.Cells worksheet and saved as .xlsx.
    class Program
    {
        static void Main()
        {
            // Prepare a DataTable containing HTML formatted text
            DataTable table = new DataTable();
            table.Columns.Add("ID");
            table.Columns.Add("Content"); // This column will hold HTML strings
            table.Columns.Add("Timestamp");

            // Add a row with HTML content (e.g., a hyperlink)
            table.Rows.Add("1", "<a href='https://www.example.com'>Example Link</a>", "2:30 PM");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure import options to treat cell values as HTML strings
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true, // Import column headers as the first row
                IsHtmlString = true      // Preserve HTML formatting during import
            };

            // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook; the HTML formatting will be retained in the cell
            workbook.Save("HtmlImportResult.xlsx");
        }
    }
}
