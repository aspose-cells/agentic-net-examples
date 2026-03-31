using System;
using Aspose.Cells;

namespace WorksheetVisibilityDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (default contains one worksheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Add additional worksheets
            workbook.Worksheets.Add("Report");
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Archive");

            // Hide the second worksheet ("Report") using the IsVisible property
            workbook.Worksheets["Report"].IsVisible = false;

            // Hide the third worksheet ("Data") using the SetVisible method (ignore errors)
            workbook.Worksheets["Data"].SetVisible(false, true);

            // Make the fourth worksheet ("Archive") very hidden (cannot be shown via UI)
            workbook.Worksheets["Archive"].VisibilityType = VisibilityType.VeryHidden;

            // Ensure the first worksheet ("Sheet1") remains visible
            workbook.Worksheets[0].IsVisible = true;

            // Save the workbook to a file
            workbook.Save("WorksheetVisibilityDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}