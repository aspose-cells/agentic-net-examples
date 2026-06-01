using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerFormulaRecalc
{
    static void Main()
    {
        // Load the template workbook that contains smart markers
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Create a sample data source (DataTable) for the smart markers
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Product");
        dt.Columns.Add("Price", typeof(double));
        dt.Rows.Add("Apple", 1.2);
        dt.Rows.Add("Banana", 0.8);

        // Initialize the WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Bind the data source to the designer
        designer.SetDataSource(dt);

        // Enable formula calculation after smart marker processing
        designer.CalculateFormula = true;

        // Process the smart markers and populate the worksheet with data
        designer.Process();

        // (Optional) Recalculate any remaining formulas in the workbook
        workbook.CalculateFormula();

        // Save the final workbook with up‑to‑date formula results
        workbook.Save("Result.xlsx");
    }
}