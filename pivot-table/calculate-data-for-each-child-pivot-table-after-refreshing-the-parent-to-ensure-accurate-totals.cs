using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotParentChildRefreshDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("Food");
            dataSheet.Cells["B3"].PutValue(150);
            dataSheet.Cells["A4"].PutValue("Drink");
            dataSheet.Cells["B4"].PutValue(80);
            dataSheet.Cells["A5"].PutValue("Drink");
            dataSheet.Cells["B5"].PutValue(120);

            // Add a worksheet for pivot tables
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Add parent pivot table
            int parentIndex = pivotSheet.PivotTables.Add("Data!A1:B5", "C3", "ParentPivot");
            PivotTable parentPivot = pivotSheet.PivotTables[parentIndex];
            parentPivot.AddFieldToArea(PivotFieldType.Row, "Category");
            parentPivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Add child pivot table using the same data source
            int childIndex = pivotSheet.PivotTables.Add("Data!A1:B5", "C15", "ChildPivot");
            PivotTable childPivot = pivotSheet.PivotTables[childIndex];
            childPivot.AddFieldToArea(PivotFieldType.Row, "Category");
            childPivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Link child pivot to parent pivot cache
            foreach (PivotField field in parentPivot.BaseFields)
            {
                childPivot.BaseFields.Add(field);
            }

            // Refresh and calculate the parent pivot table
            parentPivot.RefreshData();
            parentPivot.CalculateData();

            // Retrieve dependent (child) pivot tables
            PivotTable[] children = parentPivot.GetDependentPivotTables();

            // Refresh and calculate each child pivot table
            foreach (PivotTable child in children)
            {
                child.RefreshData();
                child.CalculateData();
            }

            // Save the workbook
            string outputPath = "ParentChildPivotRefreshDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}