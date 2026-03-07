using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class UpdateSlicerExample
{
    static void Main()
    {
        // Load the existing workbook (XLSX)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the slicer is placed on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one slicer on the sheet
        if (sheet.Slicers.Count > 0)
        {
            // Get the first slicer
            Slicer slicer = sheet.Slicers[0];

            // Example modification: change a source cell value that the slicer filters
            // (Adjust the cell reference and value as needed for your scenario)
            sheet.Cells["A2"].PutValue("UpdatedValue");

            // Refresh the slicer so it reflects the updated data and recalculates the linked PivotTable
            slicer.Refresh();
        }

        // Save the workbook with the updated slicer
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}