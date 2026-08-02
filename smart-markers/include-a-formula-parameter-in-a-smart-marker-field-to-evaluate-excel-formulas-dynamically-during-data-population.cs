// Title: C# – Use a Formula Parameter in Aspose.Cells Smart Markers for Dynamic Calculations
// Description: Demonstrates how to embed an Excel formula in a smart marker (e.g., "&=SUM(Price,Quantity)") so that Aspose.Cells evaluates the formula for each data row. The example creates a template workbook, binds a DataTable with Price and Quantity fields, enables WorkbookDesigner.CalculateFormula, processes the markers, and saves the result.
// Keywords: Aspose.Cells smart marker formula | C# WorkbookDesigner CalculateFormula | dynamic Excel calculations | smart marker SUM function | populate Excel with formulas C# | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells add formula to smart marker C# | Enable CalculateFormula for smart markers .NET | SUM smart marker example Aspose.Cells | How to compute totals with smart markers | Aspose.Cells dynamic formula evaluation
// Developer Intent: Insert a formula parameter into a smart marker so the formula is calculated automatically for every generated row during data population.
// Use Cases: Generate a sales ledger where each line total (Price × Quantity) is computed via a smart‑marker formula. | Create invoice worksheets that automatically calculate item totals without extra code. | Build summary reports that include on‑the‑fly calculations such as subtotals, averages, or custom expressions using smart markers.
// AI Prompts: Provide C# code that adds a smart marker with an Excel formula and enables CalculateFormula in Aspose.Cells. | Explain the role of WorkbookDesigner.CalculateFormula when processing smart markers containing formulas. | Show an example of binding a DataTable to a workbook and using a SUM smart marker to compute a Total column.

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates how to embed an Excel formula in a smart marker (e.g., "&=SUM(Price,Quantity)") so that Aspose.Cells evaluates the formula for each data row. The example creates a template workbook, binds a DataTable with Price and Quantity fields, enables WorkbookDesigner.CalculateFormula, processes the markers, and saves the result.
class SmartMarkerFormulaDemo
{
    static void Main()
    {
        // Create a new workbook that will serve as the template
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Price");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Total");

        // Insert smart markers for data rows
        // Simple value markers
        sheet.Cells["A2"].PutValue("&=Price");
        sheet.Cells["B2"].PutValue("&=Quantity");
        // Smart marker that contains a formula – it will be evaluated for each data row
        sheet.Cells["C2"].PutValue("&=SUM(Price,Quantity)");

        // Define the range that contains the smart markers (optional but recommended)
        sheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

        // Prepare a data source with the fields referenced by the smart markers
        DataTable data = new DataTable("Data");
        data.Columns.Add("Price", typeof(double));
        data.Columns.Add("Quantity", typeof(double));
        data.Rows.Add(10.0, 2.0);
        data.Rows.Add(15.0, 3.0);
        data.Rows.Add(20.0, 4.0);

        // Set up the WorkbookDesigner, bind the data source and enable formula calculation
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(data);
        designer.CalculateFormula = true; // ensures formulas are calculated after smart marker processing

        // Process the smart markers – this will populate the rows and evaluate the formula in column C
        designer.Process();

        // Save the populated workbook
        workbook.Save("SmartMarkerFormulaResult.xlsx");
    }
}
