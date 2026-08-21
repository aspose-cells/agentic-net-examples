// Title: Create a Pivot Table from Smart‑Marker Populated Financial Data using Aspose.Cells for .NET (C#)
// Description: C# example that builds a workbook, inserts a smart‑marker linked to a financial DataTable, processes it with WorkbookDesigner, then adds a pivot table on a new sheet (Region rows, Product columns, summed Amount) and saves the file.
// Keywords: Aspose.Cells | C# | .NET | smart markers | WorkbookDesigner | pivot table | financial data | DataTable | region product amount | programmatic Excel automation
// Common Searches: Aspose.Cells create pivot table after smart markers | C# smart marker to pivot table example | How to use WorkbookDesigner with pivot tables | Generate financial pivot report with Aspose.Cells | Programmatically add pivot table in .NET Excel
// Developer Intent: Create a pivot table that summarizes financial data inserted via smart markers.
// Use Cases: Automatically generate regional sales summaries after importing transaction data through smart markers. | Build a dynamic financial report that pivots revenue by product without manual Excel steps. | Integrate data import and analysis in a single Aspose.Cells workflow for quarterly dashboards.
// AI Prompts: Show C# code to add a filter on the Region field in the generated pivot table. | Demonstrate how to format the Amount values as currency and apply a built‑in style to the pivot table. | Explain how to export the workbook containing the pivot table to PDF using Aspose.Cells. | Provide steps to refresh the pivot table after updating the underlying smart‑marker data.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotAfterSmartMarkers
{
    // C# example that builds a workbook, inserts a smart‑marker linked to a financial DataTable, processes it with WorkbookDesigner, then adds a pivot table on a new sheet (Region rows, Product columns, summed Amount) and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // 2. Add a worksheet that will hold the smart‑marker template
                Worksheet templateSheet = workbook.Worksheets[0];
                templateSheet.Name = "Template";

                // 3. Place a smart‑marker that will be replaced by the data source.
                //    The marker syntax "&=$FinancialData" tells the designer to insert the data source named "FinancialData".
                templateSheet.Cells["A1"].PutValue("&=$FinancialData");

                // 4. Define the range that contains smart markers and give it the required name.
                //    Aspose.Cells looks for a named range called "_CellsSmartMarkers".
                AsposeRange smRange = templateSheet.Cells.CreateRange("A1");
                smRange.Name = "_CellsSmartMarkers";

                // 5. Prepare a DataTable that represents the financial data to be inserted.
                DataTable dt = new DataTable("FinancialData");
                dt.Columns.Add("Region", typeof(string));
                dt.Columns.Add("Product", typeof(string));
                dt.Columns.Add("Amount", typeof(double));

                // Sample rows
                dt.Rows.Add("North America", "Laptop", 125000);
                dt.Rows.Add("North America", "Tablet", 85000);
                dt.Rows.Add("Europe", "Laptop", 97000);
                dt.Rows.Add("Europe", "Smartphone", 66000);
                dt.Rows.Add("Asia", "Tablet", 72000);
                dt.Rows.Add("Asia", "Smartphone", 54000);

                // 6. Process smart markers using WorkbookDesigner.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("FinancialData", dt);
                designer.Process(); // processes all smart markers in the workbook

                // 7. Use the processed "Template" sheet as the source for the pivot table.
                Worksheet dataSheet = workbook.Worksheets["Template"];

                // 8. Add a new worksheet that will contain the pivot table.
                Worksheet pivotSheet = workbook.Worksheets.Add("FinancialPivot");

                // 9. Determine the source data range for the pivot table.
                //    MaxDisplayRange gives the used range of the data sheet.
                AsposeRange maxRange = dataSheet.Cells.MaxDisplayRange;
                int startRow = maxRange.FirstRow;
                int startColumn = maxRange.FirstColumn;
                int endRow = startRow + maxRange.RowCount - 1;
                int endColumn = startColumn + maxRange.ColumnCount - 1;

                string startCell = CellsHelper.CellIndexToName(startRow, startColumn);
                string endCell = CellsHelper.CellIndexToName(endRow, endColumn);
                string sourceData = $"=Template!{startCell}:{endCell}";

                // 10. Add the pivot table to the pivot sheet (lifecycle: add via PivotTableCollection.Add)
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "FinancialSummary");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // 11. Configure the pivot table fields:
                //     - Region as Row field
                //     - Product as Column field
                //     - Amount as Data field (sum)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // 12. Optional: display the pivot in tabular form and calculate the data.
                pivotTable.ShowInTabularForm();
                pivotTable.CalculateData();

                // 13. Save the workbook (lifecycle: save)
                workbook.Save("FinancialPivotAfterSmartMarkers.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
