using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a workbook and add sample data with a pivot table
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells cells = sourceSheet.Cells;

            // Populate source data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Food");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Clothing");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Electronics");
            cells["B4"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = sourceSheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivotTable = sourceSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

            // Step 2: Save the workbook to a memory stream (XLS format)
            MemoryStream memoryStream = sourceWorkbook.SaveToStream();

            // Reset the stream position before reading
            memoryStream.Position = 0;

            // Step 3: Load the workbook from the memory stream for manipulation
            Workbook loadedWorkbook = new Workbook(memoryStream);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Step 4: Manipulate pivot tables (e.g., refresh and add a new row field)
            PivotTableCollection pivots = loadedSheet.PivotTables;
            if (pivots.Count > 0)
            {
                PivotTable loadedPivot = pivots[0];

                // Refresh the pivot table to ensure it reflects any source changes
                loadedPivot.RefreshData();
                loadedPivot.CalculateData();

                // Example: Add a new row field (if there were more columns)
                // Here we just demonstrate setting RefreshDataOnOpeningFile property
                loadedPivot.RefreshDataOnOpeningFile = false;
            }

            // Step 5: Save the modified workbook to a file
            loadedWorkbook.Save("ModifiedPivotWorkbook.xlsx", SaveFormat.Xlsx);

            // Clean up
            memoryStream.Dispose();
            sourceWorkbook.Dispose();
            loadedWorkbook.Dispose();

            Console.WriteLine("Workbook loaded from memory stream, pivot table manipulated, and saved as ModifiedPivotWorkbook.xlsx");
        }
    }
}