// Title: Import HTML Table as a Structured Excel ListObject with Aspose.Cells for .NET
// Description: Demonstrates how to load an HTML file using Aspose.Cells, pick the first HTML table, convert it into an Excel ListObject (structured table) by enabling TableToListObject, and save the result as an XLSX workbook.
// Keywords: Aspose.Cells HTML table import | C# convert HTML to Excel ListObject | HtmlLoadOptions TableToListObject | load HTML into workbook Aspose | structured Excel table from HTML | .NET Aspose.Cells table load options | HTML to XLSX conversion C# | Aspose.Cells table extraction | Excel ListObject creation programmatically | Aspose.Cells example HTML table
// Common Searches: Aspose.Cells import first HTML table as ListObject | C# convert HTML table to Excel structured table | How to use HtmlLoadOptions TableToListObject | Load HTML file into workbook with Aspose.Cells | Create Excel table from HTML using .NET
// Developer Intent: Load an HTML document, transform its initial table into an Excel ListObject (native table), and write the workbook to an XLSX file.
// Use Cases: Migrate web‑based reports into Excel while preserving native table features for pivot‑tables and filtering. | Automate extraction of data tables from legacy HTML pages into structured worksheets for downstream analytics. | Generate Excel dashboards that retain the original column headings and table layout from source HTML.
// AI Prompts: Show how to import several HTML tables as separate ListObjects in one workbook using Aspose.Cells. | Explain how to rename or map column headers when converting an HTML table to a ListObject. | Provide robust error handling for missing files, empty tables, or malformed HTML with HtmlLoadOptions.

using System;
using Aspose.Cells;

namespace HtmlTableImportExample
{
    // Demonstrates how to load an HTML file using Aspose.Cells, pick the first HTML table, convert it into an Excel ListObject (structured table) by enabling TableToListObject, and save the result as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains the table
            string htmlFilePath = "input.html";

            // Path where the resulting Excel file will be saved
            string excelFilePath = "output.xlsx";

            // Create HTML load options
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Configure a table load option:
            // - TableIndex = 0  (first table in the HTML)
            // - TableToListObject = true  (convert the HTML table to an Excel ListObject, i.e., a structured table)
            HtmlTableLoadOption tableOption = new HtmlTableLoadOption
            {
                TableIndex = 0,
                TableToListObject = true
            };

            // Add the table option to the collection
            loadOptions.TableLoadOptions.Add(tableOption);

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as an Excel file
            workbook.Save(excelFilePath);

            Console.WriteLine($"HTML table imported and saved as structured table in '{excelFilePath}'.");
        }
    }
}
