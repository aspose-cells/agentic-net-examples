using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetVisibilityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add two more worksheets
            workbook.Worksheets.Add("HiddenSheet");
            workbook.Worksheets.Add("VisibleSheet");

            // Hide the second worksheet (index 1) using the IsVisible property
            workbook.Worksheets[1].IsVisible = false;

            // Save the workbook with the hidden sheet
            workbook.Save("Workbook_With_Hidden_Sheet.xlsx");

            // Verify visibility status
            Console.WriteLine("Initial visibility status:");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"{workbook.Worksheets[i].Name}: {(workbook.Worksheets[i].IsVisible ? "Visible" : "Hidden")}");
            }

            // Show the previously hidden worksheet
            workbook.Worksheets[1].IsVisible = true;

            // Save the workbook after making the sheet visible
            workbook.Save("Workbook_With_All_Sheets_Visible.xlsx");

            // Verify updated visibility status
            Console.WriteLine("\nAfter showing the hidden sheet:");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"{workbook.Worksheets[i].Name}: {(workbook.Worksheets[i].IsVisible ? "Visible" : "Hidden")}");
            }
        }
    }
}