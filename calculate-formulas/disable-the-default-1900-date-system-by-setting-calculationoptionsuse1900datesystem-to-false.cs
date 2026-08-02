// Title: Enable 1904 Date System in Aspose.Cells (C#) and Recalculate Formulas
// Description: Shows how to set Workbook.Settings.Date1904 to true, which disables the default 1900 date system, runs CalculateFormula with the 1904 epoch, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells 1904 date system | disable 1900 date system | Workbook.Settings.Date1904 | CalculateFormula C# | Excel date epoch Aspose | cross‑platform Excel dates | C# Aspose.Cells example
// Common Searches: Aspose.Cells enable 1904 date system C# | turn off 1900 date system in Aspose.Cells | Workbook.Settings.Date1904 true example | calculate formulas after changing date system Aspose | Excel 1904 epoch Aspose.Cells
// Developer Intent: Switch a workbook to the 1904 date system and recalculate its formulas.
// Use Cases: Create Excel files that match the Mac 1904 date system for cross‑platform compatibility. | Run date‑sensitive formulas without manual adjustments after changing the epoch. | Export workbooks to environments that expect the 1904 date system (e.g., legacy Mac Excel).
// AI Prompts: Provide C# code that sets Workbook.Settings.Date1904 to true and calls CalculateFormula in Aspose.Cells. | Explain the impact of switching from the 1900 to the 1904 date system on Excel formulas when using Aspose.Cells. | Show how to verify that the 1904 date system is active after changing Workbook.Settings.Date1904.

using System;
using Aspose.Cells;

// Shows how to set Workbook.Settings.Date1904 to true, which disables the default 1900 date system, runs CalculateFormula with the 1904 epoch, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable the 1904 date system (disables the default 1900 date system)
            workbook.Settings.Date1904 = true;

            // Perform calculation using the workbook's current settings (including 1904 date system)
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
