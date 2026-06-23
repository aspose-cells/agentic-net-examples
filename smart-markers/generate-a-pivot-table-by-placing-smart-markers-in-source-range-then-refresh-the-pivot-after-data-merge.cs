using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSmartMarkerPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare the source worksheet with smart markers
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Header row
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Amount");

            // Smart markers – they will be replaced by data from the DataTable
            sourceSheet.Cells["A2"].PutValue("&=Data!Category");
            sourceSheet.Cells["B2"].PutValue("&=Data!Amount");

            // -------------------------------------------------
            // 2. Create a DataTable that will be merged via smart markers
            // -------------------------------------------------
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            dt.Rows.Add("Fruits", 1200);
            dt.Rows.Add("Vegetables", 850);
            dt.Rows.Add("Beverages", 430);
            dt.Rows.Add("Snacks", 670);

            // -------------------------------------------------
            // 3. Process smart markers using WorkbookDesigner
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // Merges data into the source sheet

            // -------------------------------------------------
            // 4. Add a new worksheet for the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

            // Determine the data range after smart marker processing
            // (Assumes data starts at A1 and occupies contiguous rows/columns)
            string sourceDataRange = $"=SourceData!{sourceSheet.Cells.MaxDisplayRange.Address}";

            // Add a pivot table using the processed data range
            int pivotIndex = pivotSheet.PivotTables.Add(sourceDataRange, "A3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // -------------------------------------------------
            // 5. Refresh the pivot table to reflect merged data
            // -------------------------------------------------
            pivotSheet.RefreshPivotTables();

            // -------------------------------------------------
            // 6. Save the workbook
            // -------------------------------------------------
            workbook.Save("SmartMarkerPivotDemo.xlsx");
        }
    }
}