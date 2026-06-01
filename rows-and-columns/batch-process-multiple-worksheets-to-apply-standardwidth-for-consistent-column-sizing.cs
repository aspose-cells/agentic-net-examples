using System;
using Aspose.Cells;

class BatchStandardWidth
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add extra worksheets for demonstration purposes
        workbook.Worksheets.Add();
        workbook.Worksheets.Add();

        // Desired default column width (in character units)
        double standardWidth = 18.25;

        // Apply the same StandardWidth to every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Cells.StandardWidth gets or sets the default column width for the worksheet
            sheet.Cells.StandardWidth = standardWidth;
        }

        // Save the workbook with the updated column settings
        workbook.Save("BatchStandardWidth.xlsx");
    }
}