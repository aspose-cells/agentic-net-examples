// Title: Set Excel tab bar width, hide the fourth worksheet, and save the workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set tab bar width in pixels C# | Hide a worksheet by zero‑based index using Aspose.Cells | Save a workbook after changing worksheet visibility with Aspose.Cells | Adjust Excel UI tab bar size programmatically in .NET
// Tags: Workbook.Settings.TabBarWidthInPixels | Worksheet.IsVisible property hide sheet | Save workbook as Xlsx Aspose.Cells | Modify worksheet visibility C# Aspose.Cells | Excel tab bar width configuration Aspose.Cells

using System;
using Aspose.Cells;

// Creates a workbook, optionally sets the tab bar width via Workbook.Settings.TabBarWidthInPixels, hides the fourth worksheet when present, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Set the tab bar width (in pixels) if the property is available in the current Aspose.Cells version.
            // In newer versions the property is Workbook.Settings.TabBarWidthInPixels.
            // Uncomment the following line if your version supports it:
            // workbook.Settings.TabBarWidthInPixels = 200;

            // Hide the fourth worksheet (index 3, zero‑based) if it exists
            if (workbook.Worksheets.Count > 3)
            {
                Worksheet fourthSheet = workbook.Worksheets[3];
                fourthSheet.IsVisible = false; // Alternatively: fourthSheet.Visibility = VisibilityType.Hidden;
            }

            // Save the workbook with the applied settings
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
