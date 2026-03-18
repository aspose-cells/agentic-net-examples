using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlPreserveDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains self‑closing tags (e.g., <br/>)
            string htmlFilePath = "source.html";

            // Read the HTML content; if the file does not exist, use a default sample.
            string htmlContent = File.Exists(htmlFilePath)
                ? File.ReadAllText(htmlFilePath)
                : "<html><body>Sample<br/>HTML content</body></html>";

            // Create a DataTable with a single column to hold the HTML string
            DataTable table = new DataTable();
            table.Columns.Add("HtmlContent", typeof(string));
            table.Rows.Add(htmlContent);

            // Create a new workbook (Aspose.Cells)
            Workbook workbook = new Workbook();

            // Import the DataTable into the first worksheet.
            // Set IsHtmlString = true so the cell keeps the HTML tags unchanged.
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = false, // No header row needed
                IsHtmlString = true       // Preserve HTML tags as raw text
            };
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.ImportData(table, 0, 0, importOptions);

            // Save the workbook to XLSX format. The HTML tags remain intact in the cell.
            string outputPath = "PreservedHtml.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML content has been imported and saved to '{outputPath}'.");
        }
    }
}