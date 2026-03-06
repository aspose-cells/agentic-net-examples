using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the slicer collection on the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Ensure there is at least one slicer to format
        if (slicers.Count > 0)
        {
            // Retrieve the first slicer
            Slicer slicer = slicers[0];

            // Apply formatting to the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleDark2; // Change built‑in style
            slicer.Caption = "Formatted Slicer";                // Set a custom caption
            slicer.ShowCaption = true;                         // Ensure the caption is visible
            slicer.NumberOfColumns = 2;                        // Display items in two columns
            slicer.WidthPixel = 150;                           // Set width in pixels
            slicer.HeightPixel = 200;                          // Set height in pixels

            // Refresh the slicer to apply changes to the underlying pivot table
            slicer.Refresh();
        }

        // Save the modified workbook
        workbook.Save("output_formatted.xlsx");
    }
}