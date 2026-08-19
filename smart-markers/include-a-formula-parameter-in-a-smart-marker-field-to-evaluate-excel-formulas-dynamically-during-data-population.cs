// Title: Use a Formula Parameter in Aspose.Cells Smart Markers (C#) for Dynamic Excel Calculations
// Description: Demonstrates how to embed a formula smart marker in an Excel template, bind a DataTable, enable WorkbookDesigner.CalculateFormula, and generate a workbook where totals are computed on‑the‑fly using the PRODUCT function.
// Keywords: Aspose.Cells | smart markers | formula smart marker | C# | .NET | WorkbookDesigner | CalculateFormula | PRODUCT function | dynamic Excel formula | data binding | Excel automation
// Common Searches: Aspose.Cells formula smart marker example | how to calculate formulas after smart marker processing C# | use PRODUCT function in Aspose.Cells smart marker | enable CalculateFormula property in WorkbookDesigner | dynamic total column with smart markers
// Developer Intent: Add a Formula parameter to a smart marker so Excel formulas are evaluated automatically during data population.
// Use Cases: Generate a sales report where the Total column is calculated per row with a PRODUCT formula smart marker. | Create an invoice workbook that auto‑computes line‑item totals by multiplying unit price and quantity. | Build an inventory sheet that derives stock value for each item through a formula smart marker and saves the result without manual recalculation.
// AI Prompts: Show how to define a formula smart marker that multiplies price and quantity using Aspose.Cells for .NET. | Explain how to enable automatic formula calculation after processing smart markers with WorkbookDesigner in C#. | Provide a complete C# example that uses the PRODUCT function inside a smart marker and saves the workbook with calculated results.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    // Demonstrates how to embed a formula smart marker in an Excel template, bind a DataTable, enable WorkbookDesigner.CalculateFormula, and generate a workbook where totals are computed on‑the‑fly using the PRODUCT function.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a template workbook with smart markers
            // -------------------------------------------------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["C1"].PutValue("Quantity");
            cells["D1"].PutValue("Total");

            // Data rows using smart markers
            // &=$ColumnName   -> binds the column value
            // &=Total=PRODUCT(&=$Price,&=$Quantity) -> defines a formula smart marker
            cells["A2"].PutValue("&=$Product");
            cells["B2"].PutValue("&=$Price");
            cells["C2"].PutValue("&=$Quantity");
            cells["D2"].PutValue("&=Total=PRODUCT(&=$Price,&=$Quantity)");

            // -------------------------------------------------
            // 2. Prepare the data source
            // -------------------------------------------------
            DataTable dt = new DataTable("Sales");
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("Quantity", typeof(int));

            dt.Rows.Add("Apple", 1.20, 10);
            dt.Rows.Add("Banana", 0.80, 15);
            dt.Rows.Add("Cherry", 2.50, 5);

            // -------------------------------------------------
            // 3. Process the smart markers
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(template);
            designer.SetDataSource(dt);

            // Enable formula calculation after data binding
            designer.CalculateFormula = true;

            // Process all smart markers in the workbook
            designer.Process();

            // -------------------------------------------------
            // 4. Save the result
            // -------------------------------------------------
            designer.Workbook.Save("SmartMarkerWithFormula_Output.xlsx");

            Console.WriteLine("Workbook generated successfully.");
        }
    }
}
