// Title: C# – Toggle Aspose.Cells Worksheet.ShowFormulas to display formulas for debugging
// Description: Demonstrates how to create a workbook with Aspose.Cells, assign a SUM formula to cell A1, print the calculated result, enable the ShowFormulas flag to reveal the formula text, and save the workbook as FormulaDebugDemo.xlsx. Ideal for developers who need to inspect formulas while troubleshooting complex spreadsheets.
// Keywords: Aspose.Cells ShowFormulas | C# display Excel formulas | debug Excel formulas Aspose | toggle formula view .NET | Worksheet.ShowFormulas example | Aspose.Cells debugging techniques
// Common Searches: show formulas instead of values Aspose.Cells C# | enable formula view in generated Excel file .NET | how to toggle ShowFormulas for debugging | Aspose.Cells display cell formulas | debug complex calculations with Aspose.Cells
// Developer Intent: Provide a quick way to switch between calculated values and raw formulas in a worksheet to aid debugging.
// Use Cases: Programmatically compare a cell's evaluated result with its underlying formula during development. | Create Excel reports that default to values but allow auditors to enable formula view on demand. | Log both the formula string and its result for automated testing of spreadsheet logic.
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, inserts a formula, toggles Worksheet.ShowFormulas to show the formula, and saves the file. | Explain the relationship between Worksheet.ShowFormulas, Cell.Formula, and Cell.StringValue when switching between value and formula display. | Show how to add a macro or button in an Aspose.Cells‑generated workbook that toggles ShowFormulas at runtime for end‑user debugging.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDebugDemo
{
    // Demonstrates how to create a workbook with Aspose.Cells, assign a SUM formula to cell A1, print the calculated result, enable the ShowFormulas flag to reveal the formula text, and save the workbook as FormulaDebugDemo.xlsx. Ideal for developers who need to inspect formulas while troubleshooting complex spreadsheets.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a sample formula in cell A1
            worksheet.Cells["A1"].Formula = "=SUM(1,2,3)";

            // Initially show the calculated result
            worksheet.ShowFormulas = false;
            Console.WriteLine("ShowFormulas OFF -> Cell A1 displays: " + worksheet.Cells["A1"].StringValue);

            // Enable formula view for debugging
            worksheet.ShowFormulas = true;
            Console.WriteLine("ShowFormulas ON -> Cell A1 displays: " + worksheet.Cells["A1"].StringValue);

            // Save the workbook (lifecycle: save)
            workbook.Save("FormulaDebugDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
