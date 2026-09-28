// Title: Force US English function names for formulas in Aspose.Cells workbooks using C#
// AI Prompts: Assign a US culture object to workbook.Settings.CultureInfo so that all formulas use English function names. | Add a SUM formula with English syntax to a cell after configuring the workbook locale. | Save a workbook after enforcing US English formula parsing, independent of the operating system language. | Load an existing workbook, apply the en-US culture setting, and then modify its formulas.
// Common Searches: Aspose.Cells C# enforce English function names in formulas | How to set workbook culture to en-US for formula parsing in Aspose.Cells | Override system locale for Excel formulas when using Aspose.Cells .NET | US English Excel functions with Aspose.Cells workbook Settings.CultureInfo | Force SUM function name in Aspose.Cells regardless of OS language
// Tags: Aspose.Cells set workbook culture en-US | override formula language locale | US English Excel functions Aspose.Cells | C# workbook Settings.CultureInfo usage | enforce English function names in formulas

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates (or loads) an Aspose.Cells workbook, sets workbook.Settings.CultureInfo to the US culture to guarantee that all formulas use English function names, assigns a SUM formula using the English syntax, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (replace with load if needed)
        var workbook = new Workbook(); // new Workbook("input.xlsx") to load an existing file

        // Configure formula parsing to always use US English function names
        workbook.Settings.CultureInfo = new CultureInfo("en-US");

        // Example usage: set a formula using US function name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "SUM(B1:B10)";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
