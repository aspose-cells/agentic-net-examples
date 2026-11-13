// Title: Create multiple PivotTables on separate worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a PivotTable to each worksheet in a loop | batch generate PivotTables on new sheets with Aspose.Cells .NET | programmatically create PivotTables from range A1:D20 on separate worksheets | C# example for automating multiple PivotTables using Aspose.Cells | how to use PivotTables.Add to create PivotTables on different sheets in Aspose.Cells
// Tags: Aspose.Cells PivotTables.Add method | C# batch PivotTable creation | multiple worksheets PivotTable automation | Excel workbook PivotTable generation Aspose | source range A1:D20 PivotTable Aspose.Cells | create PivotTable on new sheet C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable related classes

namespace PivotTableBatchExample
{
    // The sample program creates a new workbook, adds three data worksheets populated with sample rows, then loops through each sheet to add a corresponding PivotTable on a newly added worksheet using the PivotTables.Add method (source range A1:D20). The workbook is saved as BatchPivotTables.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Ensure three worksheets exist
                for (int i = 0; i < 3; i++)
                {
                    if (i >= workbook.Worksheets.Count)
                        workbook.Worksheets.Add();
                }

                // Add sample data to three worksheets
                for (int i = 0; i < 3; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    ws.Name = $"DataSheet{i + 1}";

                    // Header row
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Product");
                    ws.Cells["C1"].PutValue("Quantity");
                    ws.Cells["D1"].PutValue("Price");

                    // Data rows (A2:D20)
                    for (int row = 2; row <= 20; row++)
                    {
                        ws.Cells[row - 1, 0].PutValue($"Cat{(row % 5) + 1}");
                        ws.Cells[row - 1, 1].PutValue($"Prod{(row % 10) + 1}");
                        ws.Cells[row - 1, 2].PutValue(row * 2);          // Quantity
                        ws.Cells[row - 1, 3].PutValue(row * 1.5);        // Price
                    }
                }

                // Loop over each data worksheet and create a PivotTable
                for (int i = 0; i < 3; i++)
                {
                    Worksheet dataWs = workbook.Worksheets[i];
                    string sourceDataRange = "A1:D20";

                    // Add a new worksheet to host the PivotTable
                    Worksheet pivotWs = workbook.Worksheets.Add($"PivotSheet{i + 1}");

                    // Destination cell for the PivotTable on the pivot worksheet
                    string pivotTableLocation = "A3";

                    // Create the PivotTable (using the obsolete overload for simplicity)
                    int pivotTableIndex = pivotWs.PivotTables.Add(pivotTableLocation, sourceDataRange, $"PivotTable{i + 1}");
                    PivotTable pivotTable = pivotWs.PivotTables[pivotTableIndex];

                    // NOTE: Adding fields and styles is omitted to maintain compatibility with the current Aspose.Cells version.
                    // If needed, you can configure RowFields, ColumnFields, DataFields, and styles using the appropriate API overloads.
                }

                // Save the workbook
                workbook.Save("BatchPivotTables.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully as 'BatchPivotTables.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
