using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – set all worksheets to Normal view
    class SetAllSheetsNormalView
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") to load

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the view type to Normal view
                sheet.ViewType = ViewType.NormalView;
            }

            // Save the workbook to a file
            workbook.Save("AllSheetsNormalView.xlsx");
        }
    }
}