// Title: Export Multiple Tables to HTML with Custom TableCssId Using Aspose.Cells for .NET
// Description: Creates a workbook, adds two ListObject tables (A1:C4 and E1:G4), sets HtmlSaveOptions.TableCssId to a custom value, and saves the worksheet as HTML so each exported table receives the specified CSS class prefix.
// Keywords: Aspose.Cells | HtmlSaveOptions | TableCssId | C# | export to HTML | multiple tables | ListObject | custom CSS class | .NET | HTML report
// Common Searches: Aspose.Cells set TableCssId for HTML export | export multiple ListObject tables to HTML C# | custom CSS class for tables in Aspose.Cells HTML output | HtmlSaveOptions TableCssId example | how to add CSS prefix to tables when saving as HTML
// Developer Intent: Add two ListObject tables to a worksheet and save the workbook as HTML with a custom TableCssId applied to each table.
// Use Cases: Generate an HTML report where all tables share a common CSS class for unified styling. | Integrate exported tables into a web page and target them with external CSS rules using a known class prefix. | Automate batch export of worksheets containing several tables while preserving individual table styles and a consistent CSS identifier.
// AI Prompts: Write C# code that creates three ListObject tables and saves the workbook as HTML with HtmlSaveOptions.TableCssId set to "report-table". | Explain how HtmlSaveOptions.TableCssId modifies the HTML output and show how to reference the generated tables in an external stylesheet. | Provide a step‑by‑step tutorial for exporting a worksheet with multiple tables to HTML, ensuring each table receives a custom CSS class prefix using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCssIdDemo
{
    // Creates a workbook, adds two ListObject tables (A1:C4 and E1:G4), sets HtmlSaveOptions.TableCssId to a custom value, and saves the worksheet as HTML so each exported table receives the specified CSS class prefix.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data for the first table (A1:C4)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(85);
            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(92);
            cells["A4"].PutValue(3);
            cells["B4"].PutValue("Charlie");
            cells["C4"].PutValue(78);

            // Populate data for the second table (E1:G4)
            cells["E1"].PutValue("Product");
            cells["F1"].PutValue("Category");
            cells["G1"].PutValue("Price");
            cells["E2"].PutValue("Laptop");
            cells["F2"].PutValue("Electronics");
            cells["G2"].PutValue(1200);
            cells["E3"].PutValue("Desk");
            cells["F3"].PutValue("Furniture");
            cells["G3"].PutValue(250);
            cells["E4"].PutValue("Pen");
            cells["F4"].PutValue("Stationery");
            cells["G4"].PutValue(2);

            // Add first table (ListObject) covering A1:C4
            ListObjectCollection tables = sheet.ListObjects;
            int firstTableIndex = tables.Add(0, 0, 3, 2, true);
            ListObject firstTable = tables[firstTableIndex];
            firstTable.TableStyleName = "TableStyleMedium2"; // optional style

            // Add second table covering E1:G4
            int secondTableIndex = tables.Add(0, 4, 3, 6, true);
            ListObject secondTable = tables[secondTableIndex];
            secondTable.TableStyleName = "TableStyleMedium9"; // optional style

            // Configure HTML save options to apply a CSS class prefix to tables
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // The prefix will be added to CSS class names of table elements that have TableCssId attribute
            htmlOptions.TableCssId = "custom-table";

            // Save the workbook as HTML; each table will have the specified CSS prefix
            string outputPath = "MultipleTablesWithCssId.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with TableCssId set to '{htmlOptions.TableCssId}'.");
        }
    }
}
