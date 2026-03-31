using System;
using Aspose.Cells;

namespace AsposeCellsVisibilityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add an extra worksheet for demonstration (optional)
            workbook.Worksheets.Add("Sheet2");

            // Hide the first worksheet (index 0) from the tab view
            // SetVisible(isVisible, ignoreError) -> false to hide, true to ignore errors if any
            workbook.Worksheets[0].SetVisible(false, true);

            // Alternatively, you could use:
            // workbook.Worksheets[0].VisibilityType = VisibilityType.Hidden;
            // or
            // workbook.Worksheets[0].IsVisible = false;

            // Save the workbook (lifecycle: save)
            workbook.Save("HiddenWorksheetDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}