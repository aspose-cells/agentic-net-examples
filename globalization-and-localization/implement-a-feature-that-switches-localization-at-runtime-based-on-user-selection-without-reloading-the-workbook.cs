// Title: Change workbook localization at runtime in Aspose.Cells for .NET without reloading the file
// AI Prompts: Write a C# method that takes a Workbook instance and a culture name, sets Workbook.Settings.CultureInfo to the new CultureInfo, and then calls Workbook.CalculateFormula to apply the locale instantly. | Show how to let a user pick a language at runtime and update an Aspose.Cells workbook’s localization without reopening the file, using CultureInfo and formula recalculation.
// Common Searches: aspnet change Excel workbook language on the fly using Aspose.Cells | how to update workbook culture info without reopening the file in C# | apply new locale to Aspose.Cells workbook and refresh formulas | runtime localization of Excel file with Aspose.Cells .NET | switch Excel workbook culture to German programmatically Aspose.Cells
// Tags: Workbook.Settings.CultureInfo runtime update | Apply CultureInfo and recalculate formulas Aspose.Cells | Dynamic Excel localization C# | Switch workbook locale without reload | Runtime culture switch Aspose.Cells .NET

using Aspose.Cells;
using System.Globalization;

// A helper that assigns a new CultureInfo to Workbook.Settings.CultureInfo and invokes CalculateFormula, enabling instant culture‑specific formatting and formula evaluation without reloading the workbook.
public static class LocalizationHelper
{
    // Switches the workbook's localization at runtime without reloading.
    public static void SwitchLocalization(Workbook workbook, string cultureName)
    {
        // Apply the selected culture to the workbook.
        workbook.Settings.CultureInfo = new CultureInfo(cultureName);
        // Recalculate formulas so that culture‑specific formatting takes effect.
        workbook.CalculateFormula();
    }
}

public class Program
{
    public static void Main()
    {
        // Load the workbook (using the provided load rule).
        Workbook workbook = new Workbook("input.xlsx");

        // Example: user selects German (Germany) localization.
        LocalizationHelper.SwitchLocalization(workbook, "de-DE");

        // Save the workbook (using the provided save rule).
        workbook.Save("output.xlsx");
    }
}
