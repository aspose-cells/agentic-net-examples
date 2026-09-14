// Title: Load an XLSX workbook from disk and set its formula calculation mode to Manual using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, changes the workbook's calculation mode to Manual, and shows how to save the modified file. | Demonstrate how to configure Workbook.Settings.FormulaSettings in Aspose.Cells to disable automatic formula recalculation after loading an existing Excel workbook. | Provide a step‑by‑step C# example that loads an Excel workbook, sets CalcModeType to Manual, and optionally persists the change.
// Common Searches: Aspose.Cells C# load existing Excel file and set calculation mode to manual without recalculating formulas | How to prevent automatic formula evaluation when opening an XLSX with Aspose.Cells in .NET | Set CalcModeType.Manual after loading workbook using Aspose.Cells for .NET example
// Tags: Aspose.Cells set manual calculation mode | load xlsx workbook Aspose.Cells C# | Workbook.Settings.FormulaSettings CalcModeType.Manual | disable automatic formula recalculation Aspose.Cells | manual formula evaluation Aspose.Cells .NET

using System;
using Aspose.Cells;

// // Loads an existing XLSX file with Aspose.Cells, switches the workbook's formula calculation mode to Manual via FormulaSettings, and optionally saves the workbook to persist the change.
class Program
{
    static void Main()
    {
        // Specify the path to the existing XLSX file
        string filePath = "input.xlsx";

        // Load the workbook from the file using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Set the workbook's calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // (Optional) Save the workbook if you want to persist the change
        // workbook.Save("output.xlsx");
    }
}
