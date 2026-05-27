using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class SetAllWorksheetsToNormalView
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through all worksheets and set the view type to NormalView
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ViewType = ViewType.NormalView; // Normal view ensures consistent on‑screen display
            }

            // Save the workbook to a file
            workbook.Save("AllSheetsNormalView.xlsx", SaveFormat.Xlsx);
        }
    }
}