using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a pivot table
        worksheet.Cells["A1"].PutValue("Fruit");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(5);
        worksheet.Cells["B4"].PutValue(8);

        // Add a pivot table
        int pivotIdx = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivot = worksheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add slicers and assign explicit names
        SlicerCollection slicers = worksheet.Slicers;
        int slicerIdx1 = slicers.Add(pivot, "E5", "Fruit");
        slicers[slicerIdx1].Name = "FruitSlicer1";
        int slicerIdx2 = slicers.Add(pivot, "E7", "Fruit");
        slicers[slicerIdx2].Name = "FruitSlicer2";

        // List of slicer names to be removed
        List<string> slicerNamesToRemove = new List<string>
        {
            "FruitSlicer1",
            "NonExistingSlicer",
            "FruitSlicer2"
        };

        // Process each name: attempt removal and log the result
        foreach (string name in slicerNamesToRemove)
        {
            try
            {
                // Retrieve slicer by name; throws if not found
                Slicer slicer = slicers[name];
                slicers.Remove(slicer);
                Console.WriteLine($"Removed slicer '{name}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to remove slicer '{name}': {ex.Message}");
            }
        }

        // Save the workbook with the remaining slicers (if any)
        workbook.Save("SlicersRemoved.xlsx");
    }
}