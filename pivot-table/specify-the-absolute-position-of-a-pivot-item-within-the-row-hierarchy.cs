using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotItemAbsolutePositionDemo
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
            const string sourceFile = "PivotSource.xlsx";
            const string outputFile = "PivotTable_With_ItemPositions.xlsx";

            // Verify source workbook exists
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file \"{sourceFile}\" not found.");
                return;
            }

            // Load workbook containing source data
            Workbook workbook = new Workbook(sourceFile);

            // Add a worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTableSheet");

            // Reference the worksheet that holds the source data
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
            {
                Console.WriteLine("Worksheet \"Data\" not found in the source workbook.");
                return;
            }

            // Create a pivot table based on the source range
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int ptIndex = pivotTables.Add("='Data'!A1:D500", "A3", "MyPivotTable");
            PivotTable pivotTable = pivotTables[ptIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Populate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Access the PivotItem collection of the SubCategory row field
            PivotItemCollection subCategoryItems = pivotTable.RowFields["SubCategory"]?.PivotItems;
            if (subCategoryItems == null)
            {
                Console.WriteLine("Row field \"SubCategory\" not found.");
                return;
            }

            // Set absolute positions for specific items (1‑based index)
            if (subCategoryItems["ItemA"] != null)
                subCategoryItems["ItemA"].PositionInSameParentNode = 1;

            if (subCategoryItems["ItemB"] != null)
                subCategoryItems["ItemB"].PositionInSameParentNode = 2;

            // Recalculate after changing positions
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as \"{outputFile}\".");
        }
    }
}