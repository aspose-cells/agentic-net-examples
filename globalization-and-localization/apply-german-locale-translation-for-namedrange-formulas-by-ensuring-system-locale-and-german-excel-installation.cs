// Title: How to apply German (de‑DE) locale to an Excel workbook and translate named‑range formulas with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells example for setting workbook CultureInfo to de‑DE and recalculating formulas | How to localize named‑range formulas to German in an Excel file using Aspose.Cells | Apply German locale to an Excel workbook and update formulas with Aspose.Cells .NET | Set thread and workbook culture to de‑DE before saving Excel with Aspose.Cells | Translate Excel formulas to German locale programmatically with Aspose.Cells
// Tags: set workbook CultureInfo Aspose.Cells | translate named range formulas de-DE | recalculate workbook formulas C# | German locale Excel Aspose.Cells | apply de-DE culture to Excel file

using System.Globalization;
using System.Threading;
using Aspose.Cells;

// The sample sets the .NET thread culture and the workbook's CultureInfo to German (de‑DE), loads an Excel workbook, recalculates all formulas so that named‑range expressions are translated to the German locale, and saves the updated file.
class Program
{
    static void Main()
    {
        // Set the .NET thread culture to German (de-DE)
        Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
        Thread.CurrentThread.CurrentUICulture = new CultureInfo("de-DE");

        // Load the workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx"); // {LoadWorkbook}

        // Apply German locale to the workbook so that named‑range formulas are translated
        workbook.Settings.CultureInfo = new CultureInfo("de-DE");

        // Recalculate formulas to reflect locale changes
        workbook.CalculateFormula();

        // Save the workbook (replace with desired output path)
        workbook.Save("output.xlsx"); // {SaveWorkbook}
    }
}
