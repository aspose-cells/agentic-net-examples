// Title: Export Multiple Worksheet Tables to HTML with a Custom TableCssId Prefix – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add two ListObject tables, assign built‑in table styles, set HtmlSaveOptions.TableCssId to a custom prefix, and save the sheet as HTML so each rendered table receives the specified CSS class identifier.
// Keywords: Aspose.Cells export HTML multiple tables | TableCssId Aspose.Cells | HtmlSaveOptions custom CSS class | C# ListObject to HTML | Aspose.Cells table styling | HTML output with CSS prefix | .NET workbook to HTML
// Common Searches: Aspose.Cells set TableCssId for HTML export | export worksheet with several tables to HTML C# | how to add CSS class prefix to tables in Aspose.Cells | multiple ListObjects HTMLSaveOptions example | custom table CSS identifier Aspose.Cells .NET
// Developer Intent: Generate an HTML file where every worksheet table is rendered with a user‑defined CSS class prefix.
// Use Cases: Build a web report containing two data tables that share a common CSS class for centralized styling. | Create HTML dashboards where tables need predictable identifiers for external stylesheet rules. | Integrate Aspose.Cells into a .NET web app to output worksheet tables with prefixed CSS classes for responsive design.
// AI Prompts: Write C# code that adds three ListObject tables to a worksheet and exports them to HTML with TableCssId set to "report-table". | Explain the purpose of HtmlSaveOptions.TableCssId and how it affects the HTML output of multiple tables in Aspose.Cells. | Provide a step‑by‑step tutorial for customizing the CSS class of exported tables using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook, add two ListObject tables, assign built‑in table styles, set HtmlSaveOptions.TableCssId to a custom prefix, and save the sheet as HTML so each rendered table receives the specified CSS class identifier.
class ExportMultipleTablesWithCssId
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate data for the first table (range A1:C4)
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");
        for (int row = 2; row <= 4; row++)
        {
            cells[$"A{row}"].PutValue(row - 1);
            cells[$"B{row}"].PutValue($"Person {row - 1}");
            cells[$"C{row}"].PutValue(50 + row * 5);
        }

        // Populate data for the second table (range E1:G3)
        cells["E1"].PutValue("Product");
        cells["F1"].PutValue("Qty");
        cells["G1"].PutValue("Price");
        cells["E2"].PutValue("Apple");
        cells["F2"].PutValue(10);
        cells["G2"].PutValue(0.5);
        cells["E3"].PutValue("Banana");
        cells["F3"].PutValue(20);
        cells["G3"].PutValue(0.3);

        // Add the first table to the worksheet
        ListObjectCollection tables = sheet.ListObjects;
        int firstTableIdx = tables.Add(0, 0, 3, 2, true);
        ListObject firstTable = tables[firstTableIdx];
        firstTable.TableStyleName = "TableStyleMedium2";

        // Add the second table to the worksheet
        int secondTableIdx = tables.Add(0, 4, 2, 6, true);
        ListObject secondTable = tables[secondTableIdx];
        secondTable.TableStyleName = "TableStyleMedium9";

        // Configure HTML save options to apply a CSS class prefix to each table
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.TableCssId = "custom-table";

        // Save the workbook as HTML; each table will have the specified CSS prefix
        workbook.Save("MultipleTables.html", saveOptions);
    }
}
