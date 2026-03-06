using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the slicer
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one slicer on the sheet
        if (sheet.Slicers.Count > 0)
        {
            // Get the first slicer (adjust index as needed)
            Slicer slicer = sheet.Slicers[0];

            // Locate a pivot table to disconnect from the slicer.
            // Here we use the first pivot table on the same worksheet.
            if (sheet.PivotTables.Count > 0)
            {
                PivotTable pivot = sheet.PivotTables[0];

                // Disassociate the slicer from the pivot table
                slicer.RemovePivotConnection(pivot);
            }
        }

        // Save the workbook after the disassociation
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}