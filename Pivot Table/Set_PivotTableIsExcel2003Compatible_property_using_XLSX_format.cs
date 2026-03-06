using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableExcel2003CompatibilityDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Description";
            dataSheet.Cells["C1"].Value = "Status";
            dataSheet.Cells["D1"].Value = "Date";

            // Sample rows
            dataSheet.Cells["A2"].Value = "Product1";
            dataSheet.Cells["B2"].Value = "Short description";
            dataSheet.Cells["C2"].Value = "Active";
            dataSheet.Cells["D2"].Value = DateTime.Now;

            dataSheet.Cells["A3"].Value = "Product2";
            dataSheet.Cells["B3"].Value = "Very long description that exceeds Excel 2003 limits when used in pivot tables";
            dataSheet.Cells["C3"].Value = "Inactive";
            dataSheet.Cells["D3"].Value = DateTime.Now.AddDays(-1);

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create a pivot table based on the data range A1:D3
            int pivotIndex = pivotSheet.PivotTables.Add("A1:D3", "A4", "PivotTable");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Description column

            // Set Excel 2003 compatibility mode (false allows longer text)
            pivotTable.IsExcel2003Compatible = false;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTableExcel2003CompatibilityDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            PivotTableExcel2003CompatibilityDemo.Run();
        }
    }
}