// Title: C# – Repeat Columns A‑B on Every Printed Page with Aspose.Cells
// Description: Shows how to assign PageSetup.PrintTitleColumns = "$A:$B" so that columns A and B appear as repeating titles on each printed page. The sample adds data, saves the file as PrintTitleColumnsAB.xlsx, and illustrates the effect.
// Keywords: Aspose.Cells | C# | PrintTitleColumns | repeat columns on printed pages | Excel print titles | page setup | worksheet export | Aspose.Cells .NET example | Excel repeat header columns | print layout
// Common Searches: Aspose.Cells set print title columns C# | repeat columns A B on each printed page Aspose.Cells | how to use PrintTitleColumns property | C# code to repeat header columns when printing Excel | Aspose.Cells page setup repeat columns
// Developer Intent: Set columns A and B to repeat as titles on each printed page of the workbook.
// Use Cases: Create printable reports where the first two columns act as persistent identifiers on every page. | Generate multi‑page invoices or data sheets that keep column headers visible in hard‑copy form. | Automate Excel exports that require consistent column titles for compliance documentation. | Produce catalogues or inventories where row data spans several printed pages while header columns stay fixed.
// AI Prompts: Provide C# code to set non‑adjacent columns as print titles using Aspose.Cells. | Explain how to configure both PrintTitleRows and PrintTitleColumns together and adjust margins. | Show how to programmatically verify that PrintTitleColumns is applied before printing or previewing the workbook.

using System;
using Aspose.Cells;

// Shows how to assign PageSetup.PrintTitleColumns = "$A:$B" so that columns A and B appear as repeating titles on each printed page. The sample adds data, saves the file as PrintTitleColumnsAB.xlsx, and illustrates the effect.
class SetPrintTitleColumns
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set columns A and B to repeat on each printed page
        worksheet.PageSetup.PrintTitleColumns = "$A:$B";

        // Add sample data (optional, just to illustrate the effect)
        for (int i = 0; i < 100; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1} - Column A");
            worksheet.Cells[i, 1].PutValue($"Row {i + 1} - Column B");
        }

        // Save the workbook
        workbook.Save("PrintTitleColumnsAB.xlsx");
    }
}
