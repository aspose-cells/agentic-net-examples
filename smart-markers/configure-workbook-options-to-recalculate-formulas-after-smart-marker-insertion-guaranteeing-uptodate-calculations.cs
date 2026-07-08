using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerRecalcDemo
{
    static void Main()
    {
        // Load a workbook that contains smart markers
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Initialize the designer with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Create a sample data source (DataTable) for the smart markers
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Product");
        dt.Columns.Add("Price", typeof(double));
        dt.Rows.Add("Apple", 1.2);
        dt.Rows.Add("Banana", 0.8);
        dt.Rows.Add("Cherry", 2.5);

        // Bind the data source to the designer
        designer.SetDataSource(dt);

        // Enable automatic formula calculation after processing smart markers
        designer.CalculateFormula = true;

        // Process the smart markers and populate the worksheet with data
        designer.Process();

        // Optionally ensure formulas are recalculated when the workbook is saved
        workbook.Settings.FormulaSettings.CalculateOnSave = true;

        // Save the final workbook
        workbook.Save("Result.xlsx");
    }
}