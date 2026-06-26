using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class SlicerForTableExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Assume the worksheet already contains a table (ListObject)
        // Retrieve the first table in the worksheet
        if (worksheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No table found in the worksheet.");
            return;
        }
        ListObject table = worksheet.ListObjects[0];

        // Add a slicer for the first column of the table.
        // Parameters: table, column index (0‑based), destination cell for the slicer.
        int slicerIndex = worksheet.Slicers.Add(table, 0, "E1");

        // Optionally, you can access the slicer to set additional properties
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Caption = "My Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}