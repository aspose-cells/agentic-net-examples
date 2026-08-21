// Title: Clear all slicers from a PivotTable using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with a PivotTable, add slicers, and then remove every slicer by calling the Clear() method on the worksheet's SlicerCollection before saving the file.
// Keywords: Aspose.Cells clear slicers | remove all slicers .NET | SlicerCollection.Clear | delete pivot slicers programmatically | Aspose.Cells PivotTable slicer removal
// Common Searches: how to delete all slicers in Aspose.Cells C# | clear slicer collection Aspose.Cells | remove pivot table slicers programmatically | Aspose.Cells delete multiple slicers | C# remove slicers from workbook
// Developer Intent: Programmatically remove every slicer linked to a PivotTable by clearing its SlicerCollection.
// Use Cases: Strip slicers from a generated report before distribution. | Reset a template workbook to a slicer‑free state after temporary analysis. | Delete slicers when exporting to formats that do not support them.
// AI Prompts: Write C# code that adds several slicers to a PivotTable with Aspose.Cells and then removes them all using SlicerCollection.Clear. | Explain what happens when SlicerCollection.Clear is called on a worksheet containing slicers linked to multiple PivotTables. | Provide step‑by‑step instructions to delete all slicers from a workbook while keeping the PivotTable data intact.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates how to create a workbook with a PivotTable, add slicers, and then remove every slicer by calling the Clear() method on the worksheet's SlicerCollection before saving the file.
public class RemoveAllSlicersDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(5);
            sheet.Cells["B4"].PutValue(8);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Add slicers linked to the pivot table (optional, just for demonstration)
            SlicerCollection slicers = sheet.Slicers;
            slicers.Add(pivot, "E1", "Fruit");
            slicers.Add(pivot, "E5", "Fruit");

            // Remove all slicers associated with the pivot table by clearing the collection
            slicers.Clear();

            // Save the workbook to a file
            workbook.Save("RemoveAllSlicersDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        RemoveAllSlicersDemo.Run();
    }
}
