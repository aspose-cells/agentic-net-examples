// Title: How to refresh an Aspose.Cells PivotTable with calculated fields using PivotTable.CalculateData in C#
// Common Searches: C# code to recalculate Aspose.Cells pivot table after adding a calculated field | Aspose.Cells PivotTable.CalculateData method example for .NET developers | Programmatically update pivot table results with calculated fields using Aspose.Cells | Refresh pivot table data in a workbook after inserting calculated fields in C# | Aspose.Cells .NET sample for recalculating pivot table after field addition
// Tags: Aspose.Cells PivotTable.CalculateData method | C# add calculated field to pivot table | programmatic pivot table refresh .NET | update pivot table results with calculated fields | Aspose.Cells pivot table data recalculation

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example creates a workbook, fills a worksheet with sample sales data, adds a second sheet, defines a source range, and builds a pivot table with 'Category' as a row field and 'Quantity' as a data field. To ensure the pivot reflects any calculated fields, call pivotTable.CalculateData() after adding those fields before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it "Data"
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Quantity");
            dataSheet.Cells["D1"].PutValue("Price");

            string[,] sampleData = new string[,]
            {
                {"Electronics","Phone","10","500"},
                {"Electronics","Laptop","5","1200"},
                {"Furniture","Chair","20","80"},
                {"Furniture","Table","15","150"},
                {"Electronics","Tablet","8","300"},
                {"Furniture","Sofa","3","700"}
            };

            for (int i = 0; i < sampleData.GetLength(0); i++)
            {
                for (int j = 0; j < sampleData.GetLength(1); j++)
                {
                    dataSheet.Cells[i + 1, j].PutValue(sampleData[i, j]);
                }
            }

            // Add a new worksheet for the pivot table
            int pivotSheetIdx = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIdx];
            pivotSheet.Name = "Pivot";

            // Define the source data range (including header)
            int totalRows = sampleData.GetLength(0) + 1; // +1 for header row
            string sourceRange = $"Data!A1:D{totalRows}";

            // Add a pivot table at cell A3 of the pivot sheet
            int pivotIdx = pivotSheet.PivotTables.Add(sourceRange, "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];

            // Add "Category" as a row field (source field index 0)
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);

            // Add "Quantity" as a data field (source field index 2)
            pivotTable.DataFields.Add(pivotTable.DataFields[2]);

            // Save the workbook
            string outputPath = "PivotWithCalculatedField.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
