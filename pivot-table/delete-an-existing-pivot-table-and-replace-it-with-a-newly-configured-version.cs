using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotReplaceDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the worksheet that contains the pivot table (first worksheet in this example)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on the worksheet
            PivotTableCollection pivotTables = sheet.PivotTables;

            // If there is at least one pivot table, remove the first one
            if (pivotTables.Count > 0)
            {
                // Retrieve the pivot table to be removed
                PivotTable oldPivot = pivotTables[0];

                // Delete the pivot table and its data
                pivotTables.Remove(oldPivot);
            }

            // Define the source data range for the new pivot table
            // Adjust the range as needed for your data layout
            string sourceData = "A1:B4";

            // Add a new pivot table at the desired location
            int newIndex = pivotTables.Add(sourceData, "D5", "NewPivotTable");

            // Retrieve the newly added pivot table
            PivotTable newPivot = pivotTables[newIndex];

            // Configure the pivot table fields (example: Row = "Product", Data = "Sales")
            newPivot.AddFieldToArea(PivotFieldType.Row, "Product");
            newPivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to populate it with data
            newPivot.RefreshData();
            newPivot.CalculateData();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}