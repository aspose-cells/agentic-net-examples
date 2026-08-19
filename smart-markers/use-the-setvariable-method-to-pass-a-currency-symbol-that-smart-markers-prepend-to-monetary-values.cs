// Title: C# SetVariable to Prepend a Currency Symbol in Aspose.Cells Smart Markers
// Description: Demonstrates how to store a currency symbol in a worksheet variable and reference it with smart markers (e.g., "&=Variables!B1&Price") using WorkbookDesigner.SetVariable, then generate a formatted Excel file with prices prefixed by the symbol.
// Keywords: Aspose.Cells | SetVariable | C# | smart markers | currency symbol | price formatting | WorkbookDesigner | Excel template | dynamic currency | financial report | multi‑currency invoice
// Common Searches: Aspose.Cells SetVariable example C# | prepend currency symbol smart markers Aspose | how to use variables with smart markers in .NET | dynamic currency formatting Aspose.Cells | C# Excel template currency symbol SetVariable
// Developer Intent: Insert a configurable currency symbol before price values in smart‑marker‑driven Excel reports.
// Use Cases: Create a single‑source currency variable for all price columns in a quarterly financial statement. | Switch between $, €, £, or ¥ at runtime to generate locale‑specific invoices without changing the template. | Build reusable Excel templates where monetary fields automatically display the correct symbol for multi‑currency dashboards.
// AI Prompts: Generate C# code that uses WorkbookDesigner.SetVariable to set a currency symbol and applies it in smart markers for price columns. | Explain step‑by‑step how to change the currency symbol at runtime before calling Designer.Process() in Aspose.Cells. | Show an example of combining SetVariable with smart markers to produce an invoice workbook that supports multiple currencies.

using System;
using Aspose.Cells;

// Demonstrates how to store a currency symbol in a worksheet variable and reference it with smart markers (e.g., "&=Variables!B1&Price") using WorkbookDesigner.SetVariable, then generate a formatted Excel file with prices prefixed by the symbol.
class SetVariableCurrencyDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // (Optional) Add a worksheet that holds variable values
            Worksheet variablesSheet = workbook.Worksheets.Add("Variables");
            variablesSheet.Cells["A1"].PutValue("CurrencySymbol");
            variablesSheet.Cells["B1"].PutValue("$");

            // Create a WorkbookDesigner for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Add a template worksheet with smart markers
            Worksheet template = workbook.Worksheets.Add("Template");

            // Header row
            template.Cells["A1"].PutValue("Item");
            template.Cells["B1"].PutValue("Price");

            // Sample data rows (the data source for the smart marker)
            template.Cells["A2"].PutValue("Apple");
            template.Cells["B2"].PutValue(1.25);
            template.Cells["A3"].PutValue("Banana");
            template.Cells["B3"].PutValue(0.75);

            // Smart marker that prepends the currency symbol to the monetary value.
            // Use the variable stored in the Variables sheet (cell B1).
            template.Cells["C1"].PutValue("Formatted Price");
            template.Cells["C2"].PutValue("&=Variables!B1&Price");
            template.Cells["C3"].PutValue("&=Variables!B1&Price");

            // Process the smart markers
            designer.Process();

            // Save the workbook
            workbook.Save("SetVariableCurrencyDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
