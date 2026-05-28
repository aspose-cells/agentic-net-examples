using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableExcel2003CompatibilityDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Prepare data worksheet
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header row
                dataSheet.Cells["A1"].Value = "Product";
                dataSheet.Cells["B1"].Value = "Description";

                // Data rows
                dataSheet.Cells["A2"].Value = "Item1";
                dataSheet.Cells["B2"].Value = "Short description";

                // Long description (>255 chars)
                string longDescription = new string('X', 300);
                dataSheet.Cells["A3"].Value = "Item2";
                dataSheet.Cells["B3"].Value = longDescription;

                // Add a worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create pivot table from source range A1:B3, placed at A5
                int pivotIndex = pivotSheet.PivotTables.Add("A1:B3", "A5", "MyPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields: Product as row, Description as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column B

                // Enforce Excel 2003 compatibility (truncates strings >255 chars)
                pivotTable.IsExcel2003Compatible = true;

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableExcel2003CompatibilityDemo.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableExcel2003CompatibilityDemo.Run();
        }
    }
}