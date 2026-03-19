using System;
using Aspose.Cells;

namespace AsposeCellsVisibilityDemo
{
    public class WorksheetVisibilityToggle
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Add two additional worksheets
            workbook.Worksheets.Add("HiddenSheet");
            workbook.Worksheets.Add("AnotherSheet");

            // Initially hide the second worksheet using the IsVisible property
            // Sheet index 1 corresponds to "HiddenSheet"
            workbook.Worksheets[1].IsVisible = false;

            // Hide the third worksheet using the SetVisible method (ignore errors)
            // Sheet index 2 corresponds to "AnotherSheet"
            workbook.Worksheets[2].SetVisible(false, true);

            // Save the workbook with hidden sheets (lifecycle save rule)
            workbook.Save("Workbook_With_Hidden_Sheets.xlsx");

            // --- Runtime toggle: make the previously hidden sheets visible ---

            // Show the second worksheet by setting IsVisible to true
            workbook.Worksheets[1].IsVisible = true;

            // Show the third worksheet using SetVisible
            workbook.Worksheets[2].SetVisible(true, true);

            // Save the workbook after toggling visibility
            workbook.Save("Workbook_With_Visible_Sheets.xlsx");

            // Output the visibility status of each worksheet to the console
            Console.WriteLine("Worksheet visibility after toggling:");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                Console.WriteLine($"{ws.Name}: {(ws.IsVisible ? "Visible" : "Hidden")}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetVisibilityToggle.Run();
        }
    }
}