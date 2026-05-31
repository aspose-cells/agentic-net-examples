using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the merged workbook (replace with actual path)
        Workbook workbook = new Workbook("merged.xlsx");

        // Apply AutoFitRows to every worksheet to adjust row heights based on content
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.AutoFitRows();
        }

        // Save the workbook after autofitting rows
        workbook.Save("merged_autofit.xlsx");
    }
}