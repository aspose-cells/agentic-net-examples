using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one table (ListObject) in the worksheet
        if (worksheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No table found in the worksheet.");
            return;
        }

        // Get the first table in the worksheet
        ListObject table = worksheet.ListObjects[0];

        // Add a slicer for the first column of the table and place it at cell D1
        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Optional: customize slicer properties
        slicer.Caption = "Category Slicer";

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}