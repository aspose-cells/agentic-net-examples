using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace HtmlToExcelNestedTables
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains nested tables
            string htmlFilePath = "input.html";

            // Path for the generated Excel workbook
            string excelFilePath = "output.xlsx";

            // Create HtmlLoadOptions to control how HTML tables are imported
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // --------------------------------------------------------------
            // Map each HTML table (including nested ones) to a separate worksheet.
            // Table indexes are zero‑based and correspond to the order in which
            // Aspose.Cells encounters <table> elements while parsing the HTML.
            // The target sheet index specifies the worksheet that will receive the
            // table data. Here we map table 0 → sheet 0, table 1 → sheet 1, etc.
            // --------------------------------------------------------------
            loadOptions.TableLoadOptions.Add(0, 0); // First (outer) table → Sheet0
            loadOptions.TableLoadOptions.Add(1, 1); // First nested table → Sheet1
            loadOptions.TableLoadOptions.Add(2, 2); // Second nested table (if any) → Sheet2
            // Add more mappings as required for additional tables.

            // Load the HTML file into a Workbook using the configured options.
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as an Excel file.
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been converted to Excel with each table on a separate worksheet.");
        }
    }
}