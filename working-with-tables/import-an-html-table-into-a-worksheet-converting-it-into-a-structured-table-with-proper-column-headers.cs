// Title: Import an HTML Table as an Excel Structured ListObject with Aspose.Cells for .NET (C#)
// Description: C# code that uses Aspose.Cells to load an HTML file, sets HtmlLoadOptions so the first HTML table becomes an Excel ListObject (structured table), and saves the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# HTML table import | HtmlLoadOptions | HtmlTableLoadOption | ListObject | Excel structured table | convert HTML to Excel | TableToListObject | load HTML workbook .NET | Excel automation tutorial
// Common Searches: Aspose.Cells import HTML table C# | Convert HTML table to Excel ListObject .NET | HtmlLoadOptions TableToListObject example | Load HTML as structured table using Aspose.Cells | C# code to import HTML table into Excel workbook
// Developer Intent: Load an HTML document and automatically transform its first table into an Excel ListObject with Aspose.Cells.
// Use Cases: Migrate a web‑based product catalog into an Excel workbook where the data is ready for filtering, sorting, and pivot tables. | Automate the conversion of scraped HTML reports into structured Excel tables for downstream analytics. | Create a repeatable process that imports legacy HTML dashboards into Excel while preserving table semantics.
// AI Prompts: Generate C# code to import multiple HTML tables as separate ListObjects, each with a different TableIndex, using Aspose.Cells. | Show how to apply a custom style to the ListObject created from an HTML table after loading it with Aspose.Cells. | Explain strategies for handling merged cells in an HTML table when converting it to an Excel structured table with HtmlTableLoadOption.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlImportExample
{
    // C# code that uses Aspose.Cells to load an HTML file, sets HtmlLoadOptions so the first HTML table becomes an Excel ListObject (structured table), and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file containing the table
            string htmlPath = "input.html";

            // Path where the resulting Excel file will be saved
            string excelPath = "output.xlsx";

            // Create HTML load options
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Configure a table load option:
            // - TableIndex = 0 (first table in the HTML)
            // - TableToListObject = true (convert the HTML table to an Excel ListObject, i.e., a structured table)
            HtmlTableLoadOption tableOption = new HtmlTableLoadOption
            {
                TableIndex = 0,
                TableToListObject = true
            };

            // Add the table option to the collection
            loadOptions.TableLoadOptions.Add(tableOption);

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an Excel file
            workbook.Save(excelPath);

            Console.WriteLine($"HTML table imported and saved as structured table in '{excelPath}'.");
        }
    }
}
