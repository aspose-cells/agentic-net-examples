using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

public class UpdateSlicerDemo
{
    public static void Run()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the worksheet that contains the slicer (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Ensure there is at least one slicer to work with
        if (slicers.Count > 0)
        {
            // Retrieve the first slicer (or use an index you require)
            Slicer slicer = slicers[0];

            // Update slicer properties as needed
            slicer.Caption = "Updated Slicer Caption";
            slicer.NumberOfColumns = 2;          // Example: display items in two columns
            slicer.LockedPosition = false;       // Allow the user to move/resize the slicer

            // Refresh the slicer to recalculate any linked PivotTables
            slicer.Refresh();
        }

        // Save the modified workbook to a new file
        workbook.Save("Output.xlsx");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        UpdateSlicerDemo.Run();
    }
}