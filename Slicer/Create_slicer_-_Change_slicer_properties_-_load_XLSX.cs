using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (assumed to contain a pivot table)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            System.Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Retrieve the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Add a slicer linked to the pivot table field "fruit"
        // The slicer will be placed with its top‑left corner at cell E2
        int slicerIdx = worksheet.Slicers.Add(pivotTable, "E2", "fruit");
        Slicer slicer = worksheet.Slicers[slicerIdx];

        // Modify slicer properties
        slicer.Caption = "Fruit Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleDark2; // Built‑in style
        slicer.NumberOfColumns = 2;                         // Display items in two columns
        slicer.LockedPosition = true;                       // Prevent moving/resizing via UI
        slicer.ShowCaption = true;                          // Show the header caption
        slicer.WidthPixel = 200;                            // Width in pixels
        slicer.HeightPixel = 150;                           // Height in pixels

        // Refresh the slicer to ensure it reflects the current pivot data
        slicer.Refresh();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}