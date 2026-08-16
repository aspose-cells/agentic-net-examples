// Title: Load an XLSX workbook and set formula calculation mode to Manual with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an existing XLSX file using Aspose.Cells for .NET, change the workbook’s FormulaSettings.CalculationMode to Manual, and optionally save the file to keep the setting.
// Keywords: Aspose.Cells | C# | load workbook | manual calculation mode | FormulaSettings | CalcModeType.Manual | disable automatic recalculation | Excel performance
// Common Searches: Aspose.Cells set calculation mode manual C# | load XLSX file with Aspose.Cells .NET | prevent automatic formula recalculation Aspose.Cells | change workbook calculation mode programmatically | save workbook after changing calculation settings Aspose.Cells
// Developer Intent: Open an XLSX workbook and switch its formula engine to manual calculation.
// Use Cases: Boost performance when updating thousands of cells in a large workbook by disabling automatic recalculation. | Create a report template where formulas are evaluated only after all data has been inserted. | Import external data into a spreadsheet without triggering intermediate calculations, then recalculate once at the end.
// AI Prompts: Write C# code using Aspose.Cells to load an XLSX file, set CalculationMode to Manual, and save it. | Explain the difference between Manual and Automatic calculation modes in Aspose.Cells and describe scenarios for each. | Show how to batch‑update cells while the workbook is in Manual mode and then trigger a full recalculation.

using Aspose.Cells;

// Demonstrates how to open an existing XLSX file using Aspose.Cells for .NET, change the workbook’s FormulaSettings.CalculationMode to Manual, and optionally save the file to keep the setting.
class Program
{
    static void Main()
    {
        // Path to the existing XLSX file
        string filePath = "input.xlsx";

        // Load the workbook from the specified file path
        Workbook workbook = new Workbook(filePath);

        // Set the calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // (Optional) Save the workbook if you need to persist the change
        // workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
