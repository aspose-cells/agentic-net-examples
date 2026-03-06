using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the slicer collection on the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Ensure there is at least one slicer to format
        if (slicers.Count > 0)
        {
            // Retrieve the first slicer (you can also use slicers["SlicerName"] if you know the name)
            Slicer slicer = slicers[0];

            // ----- Formatting the slicer -----
            // Apply a built‑in slicer style
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Set the caption and make it visible
            slicer.Caption = "Sales Region";
            slicer.ShowCaption = true;

            // Arrange items in two columns
            slicer.NumberOfColumns = 2;

            // Adjust column width and row height (pixel units)
            slicer.ColumnWidthPixel = 120;
            slicer.RowHeightPixel = 30;

            // Control user interaction
            slicer.LockedPosition = false;   // Allow moving/resizing
            slicer.IsLocked = false;         // Allow editing
            slicer.IsPrintable = true;       // Include slicer in prints
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}