using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class UpdateSlicerExample
{
    static void Main()
    {
        // Paths to the input and output Excel files
        string inputPath = "sample.xlsx";
        string outputPath = "sample_updated.xlsx";

        // Load the existing workbook (XLSX format)
        Workbook workbook = new Workbook(inputPath);

        // Assume the slicer is placed on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access the slicer collection of the worksheet
        SlicerCollection slicers = sheet.Slicers;

        // If there is at least one slicer, refresh it (this updates its state based on the underlying PivotTable)
        if (slicers.Count > 0)
        {
            // Get the first slicer
            Slicer slicer = slicers[0];

            // Refresh the slicer to reflect any changes in the PivotTable data
            slicer.Refresh();
        }

        // Save the modified workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}