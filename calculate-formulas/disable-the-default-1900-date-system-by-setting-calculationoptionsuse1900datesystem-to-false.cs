using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable the default 1900 date system.
        // Aspose.Cells uses the 1900 system by default.
        // Setting the 1904 date system to true effectively disables the 1900 system.
        // This can be done via WorkbookSettings or via CalculationOptions.
        // Here we demonstrate both approaches.

        // Approach 1: Using WorkbookSettings
        workbook.Settings.Date1904 = true; // Enables 1904 date system (disables 1900)

        // Approach 2: Using CalculationOptions (if supported)
        CalculationOptions calcOptions = new CalculationOptions();

        // The property Use1904DateSystem (or Use1900DateSystem) controls the date system.
        // Setting it to false for Use1900DateSystem disables the 1900 system.
        // If the property exists, the following line will compile; otherwise, comment it out.
        // calcOptions.Use1900DateSystem = false; // Uncomment if the property exists

        // Alternatively, if the library provides Use1904DateSystem, set it to true.
        // calcOptions.Use1904DateSystem = true; // Uncomment if the property exists

        // Apply the calculation options (optional, e.g., for formula recalculation)
        workbook.CalculateFormula(calcOptions);

        // Save the workbook to verify the setting
        workbook.Save("DateSystemDisabled.xlsx");
    }
}