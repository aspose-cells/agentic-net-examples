// Title: ShowFormulas Property – Display Formulas Instead of Values in Aspose.Cells for .NET
// Description: Demonstrates how to toggle the Worksheet.ShowFormulas flag in Aspose.Cells (C#) to show the formula text rather than the evaluated result, useful for debugging. The example creates a workbook, assigns a formula to A1, prints the value with ShowFormulas off, enables it, prints the formula, and saves the file.
// Keywords: Aspose.Cells ShowFormulas | display Excel formulas .NET | Worksheet.ShowFormulas property | debug Excel formulas C# | toggle formula view Aspose.Cells | C# Aspose.Cells example
// Common Searches: how to view formulas in Aspose.Cells workbook | Aspose.Cells show formulas for debugging | Worksheet.ShowFormulas C# example | display formula text instead of value Aspose.Cells | toggle ShowFormulas property in .NET
// Developer Intent: Enable a worksheet to show formula strings rather than calculated results for debugging purposes.
// Use Cases: Verify that formulas are correctly written before publishing a report. | Create a debugging copy of a workbook that reveals all formulas. | Switch between formula view and value view dynamically during development.
// AI Prompts: Generate C# code that sets Worksheet.ShowFormulas to true for all sheets in a workbook and saves a debug version. | Explain the impact of Worksheet.ShowFormulas on cell.StringValue and how to retrieve the original formula. | Provide a script that toggles ShowFormulas on a specific worksheet, prints both the value and the formula, and then restores the original setting.

using System;
using Aspose.Cells;

// Demonstrates how to toggle the Worksheet.ShowFormulas flag in Aspose.Cells (C#) to show the formula text rather than the evaluated result, useful for debugging. The example creates a workbook, assigns a formula to A1, prints the value with ShowFormulas off, enables it, prints the formula, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set a formula in cell A1
        cells["A1"].Formula = "=1+2+3";

        // Show the calculated result (default behavior)
        worksheet.ShowFormulas = false;
        Console.WriteLine("ShowFormulas OFF: " + cells["A1"].StringValue);

        // Enable formula display for debugging purposes
        worksheet.ShowFormulas = true;
        Console.WriteLine("ShowFormulas ON: " + cells["A1"].StringValue);

        // Save the workbook (optional)
        workbook.Save("FormulaDebug.xlsx");
    }
}
