// Title: Use a formula smart marker to compute the Total column for each row generated from a DataTable with Aspose.Cells for .NET
// AI Prompts: Insert a smart marker that contains a formula multiplying Quantity and UnitPrice, then process the workbook with WorkbookDesigner. | Enable the RepeatFormulasWithSubtotal property so the formula is automatically copied to every row created by smart markers. | After processing, call CalculateFormula and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# how to apply a formula smart marker to each generated row | repeat formulas when expanding smart marker rows Aspose.Cells .NET | calculate total column using smart markers and DataTable in Aspose.Cells | set RepeatFormulasWithSubtotal property for formula smart markers Aspose.Cells
// Tags: formula smart marker usage Aspose.Cells | repeat formulas with subtotal Aspose.Cells | DataTable to Excel rows Aspose.Cells | auto-calculate column after smart marker processing | WorkbookDesigner repeat formulas C#

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, defines smart markers (including a formula marker in column D), enables RepeatFormulasWithSubtotal, processes a DataTable source, calculates all formulas, and saves the result as SmartMarkerFormulaDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Template with smart markers -----
        // Header row
        sheet.Cells["A1"].PutValue("ProductID");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("UnitPrice");
        sheet.Cells["D1"].PutValue("Total"); // column that will contain a formula

        // Row 2 contains smart markers.
        // &=$ColumnName inserts the value from the data source.
        // &=[Formula] inserts a formula that will be repeated for each generated row.
        sheet.Cells["A2"].PutValue("&=$ProductID");
        sheet.Cells["B2"].PutValue("&=$Quantity");
        sheet.Cells["C2"].PutValue("&=$UnitPrice");
        // Formula smart marker: total = Quantity * UnitPrice
        sheet.Cells["D2"].PutValue("&=[Quantity]*[UnitPrice]");

        // ----- Configure the designer -----
        // RepeatFormulasWithSubtotal ensures that any existing formulas are also repeated
        // when rows are inserted by the smart marker processor.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.RepeatFormulasWithSubtotal = true;

        // ----- Prepare data source -----
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductID", typeof(int));
        dt.Columns.Add("Quantity", typeof(int));
        dt.Columns.Add("UnitPrice", typeof(double));

        dt.Rows.Add(101, 2, 15.5);
        dt.Rows.Add(102, 5, 9.99);
        dt.Rows.Add(103, 1, 25.0);

        // Set the data source and process the smart markers
        designer.SetDataSource(dt);
        designer.Process();

        // Calculate all formulas so that the Total column shows actual values
        workbook.CalculateFormula();

        // Save the result
        workbook.Save("SmartMarkerFormulaDemo.xlsx");
    }
}
