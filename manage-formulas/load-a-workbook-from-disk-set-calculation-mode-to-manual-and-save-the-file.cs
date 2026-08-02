// Title: C# – Load Excel workbook, set manual formula calculation mode, and save with Aspose.Cells
// Description: Shows how to open an existing .xlsx file using Aspose.Cells for .NET, switch the workbook’s FormulaSettings.CalculationMode to Manual, and write the updated file to disk.
// Keywords: Aspose.Cells | C# | load workbook | manual calculation mode | CalcModeType.Manual | formula settings | save workbook | Excel automation | disable automatic recalculation | performance optimization
// Common Searches: Aspose.Cells set calculation mode manual | C# load Excel file and disable auto calculation | How to change formula calculation mode in Aspose.Cells | Save workbook after changing formula settings Aspose.Cells | Manual recalculation Aspose.Cells .NET
// Developer Intent: Change a workbook’s formula calculation mode to Manual and persist the setting by saving the file.
// Use Cases: Improve processing speed when performing bulk edits by turning off automatic recalculation before saving. | Distribute a spreadsheet that should not recalculate formulas on open, ensuring users trigger calculation manually. | Create a template that defers formula evaluation until a later step in a workflow.
// AI Prompts: Provide C# code that opens an Excel file with Aspose.Cells, sets CalcModeType.Manual, and saves it as a new workbook. | Explain how to disable automatic formula calculation in an existing workbook using Aspose.Cells for .NET. | Show how to verify that the calculation mode is Manual after saving a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to open an existing .xlsx file using Aspose.Cells for .NET, switch the workbook’s FormulaSettings.CalculationMode to Manual, and write the updated file to disk.
class Program
{
    static void Main()
    {
        // Load the workbook from disk
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Set calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Save the workbook (overwrites or creates a new file)
        string outputFile = "output.xlsx";
        workbook.Save(outputFile);
    }
}
