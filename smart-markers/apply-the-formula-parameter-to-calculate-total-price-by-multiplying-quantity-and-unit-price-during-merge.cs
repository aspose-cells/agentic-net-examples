// Title: Apply a Formula to Calculate Total Price with Smart Markers during Data Merge in Aspose.Cells for .NET (C#)
// Description: C# example that builds a workbook template with smart markers for Quantity and UnitPrice, sets the Total column formula (=A2*B2), merges a DataTable using WorkbookDesigner, enables RepeatFormulasWithSubtotal so the formula repeats for each generated row, recalculates all formulas, and saves the result as MergedWithTotal.xlsx.
// Keywords: Aspose.Cells | C# | Smart Markers | Excel formula repeat | RepeatFormulasWithSubtotal | WorkbookDesigner | Data merge | Calculate total price | DataTable to Excel | Invoice automation | Excel automation | Total column formula
// Common Searches: Aspose.Cells repeat formulas after smart marker merge | How to calculate total column with smart markers .NET | WorkbookDesigner RepeatFormulasWithSubtotal example | C# merge DataTable into Excel with formula | Apply formula to each row during smart marker merge
// Developer Intent: Add a multiplication formula that computes total price for each merged row and evaluate the results automatically.
// Use Cases: Generate an invoice workbook where each line‑item total is calculated on the fly during data merge. | Create a sales report that lists quantities, unit prices, and automatically derives row‑level totals without manual copying. | Build an order‑summary sheet that repeats pricing formulas for every record imported from a database or CSV file.
// AI Prompts: Show how to format the Total column as currency after the formulas are calculated. | Provide an example of using multiple smart‑marker tables in one workbook while repeating formulas for each table. | Explain the purpose of RepeatFormulasWithSubtotal and how to achieve the same outcome with manual formula replication.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsFormulaMergeDemo
{
    // C# example that builds a workbook template with smart markers for Quantity and UnitPrice, sets the Total column formula (=A2*B2), merges a DataTable using WorkbookDesigner, enables RepeatFormulasWithSubtotal so the formula repeats for each generated row, recalculates all formulas, and saves the result as MergedWithTotal.xlsx.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Set up the template with smart markers for data merge
            // Header row
            sheet.Cells["A1"].PutValue("Quantity");
            sheet.Cells["B1"].PutValue("UnitPrice");
            sheet.Cells["C1"].PutValue("Total");

            // Data row with smart markers
            sheet.Cells["A2"].PutValue("&=$Quantity");      // will be replaced by Quantity column
            sheet.Cells["B2"].PutValue("&=$UnitPrice");    // will be replaced by UnitPrice column

            // Formula cell – multiply the two columns of the same row
            // The formula uses relative references; it will be repeated for each merged row
            sheet.Cells["C2"].Formula = "=A2*B2";

            // 3. Create a DataTable that will be merged into the template
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("UnitPrice", typeof(double));

            dt.Rows.Add(2, 15.5);
            dt.Rows.Add(5, 9.99);
            dt.Rows.Add(1, 120.0);

            // 4. Use WorkbookDesigner to merge the data
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);

            // Enable repeating of the formula for each generated row
            designer.RepeatFormulasWithSubtotal = true;

            // Process the template (merge data and repeat formulas)
            designer.Process();

            // 5. Calculate all formulas so that the Total column contains the computed values
            workbook.CalculateFormula();

            // 6. Save the resulting workbook (lifecycle rule)
            workbook.Save("MergedWithTotal.xlsx");
        }
    }
}
