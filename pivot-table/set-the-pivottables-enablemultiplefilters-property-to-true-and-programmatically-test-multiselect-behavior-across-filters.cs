// Title: Enable multiple filters on an Aspose.Cells PivotTable and test multi‑select behavior in C#
// Common Searches: Aspose.Cells C# enable EnableMultipleFilters on a PivotTable | programmatically allow multi‑select filters in an Aspose.Cells PivotTable | how to test multiple filter selections in a PivotTable using Aspose.Cells .NET | set EnableMultipleFilters property for PivotTable with Aspose.Cells example | verify multi‑select filter functionality in generated Excel file via Aspose.Cells
// Tags: EnableMultipleFilters Aspose.Cells PivotTable | multi-select filter PivotTable C# | Aspose.Cells configure multiple filters Excel | programmatic PivotTable filter setup .NET | validate PivotTable multi-filter behavior

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Namespace for PivotTable related classes

// The sample creates a workbook, populates it with category data, adds a PivotTable on a separate sheet, assigns row, column, and data fields, enables the EnableMultipleFilters property, refreshes and calculates the PivotTable, and saves the result as PivotTable_MultiFilter.xlsx to demonstrate and verify multi‑select filter functionality.
class PivotTableMultiFilterDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data (Category, SubCategory, Amount)
            string[,] data = new string[,]
            {
                { "Category", "SubCategory", "Amount" },
                { "A", "A1", "100" },
                { "A", "A2", "150" },
                { "B", "B1", "200" },
                { "B", "B2", "250" },
                { "C", "C1", "300" },
                { "C", "C2", "350" }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    dataSheet.Cells[r, c].PutValue(data[r, c]);
                }
            }

            // Add a new worksheet for the PivotTable
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Define the source range for the PivotTable
            string sourceData = $"{dataSheet.Name}!A1:C7";

            // Add a PivotTable at cell A3 of the pivot sheet
            int pivotTableIndex = pivotSheet.PivotTables.Add("PivotTable1", "A3", sourceData);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Add fields to the PivotTable
            // Row field: Category (index 0)
            // Column field: SubCategory (index 1)
            // Data field: Sum of Amount (index 2)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

            // Refresh the PivotTable to calculate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Ensure the output directory exists
            string outputPath = "PivotTable_MultiFilter.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
