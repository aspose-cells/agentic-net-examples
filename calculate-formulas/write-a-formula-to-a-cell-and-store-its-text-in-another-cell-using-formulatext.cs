// Title: Aspose.Cells for .NET: Write a Formula and Capture Its Text with FORMULATEXT (C#)
// Description: This C# example creates a new workbook, assigns the formula "=SUM(1,2,3)" to cell A1, uses the Excel FORMULATEXT function in cell B1 to store the literal formula string, forces calculation so B1 displays the text, prints both values to the console, and saves the file as FormulaTextDemo.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | FORMULATEXT | FormulaText | retrieve formula as text | write formula programmatically | calculate formulas | Excel formula text | sample code | GitHub example
// Common Searches: Aspose.Cells FORMULATEXT C# example | How to get formula text with Aspose.Cells | Write formula and retrieve its string in .NET | Aspose.Cells calculate formulas programmatically | Store Excel formula as text using Aspose
// Developer Intent: Capture the textual representation of a cell's formula in another cell using FORMULATEXT.
// Use Cases: Audit worksheets that show both results and the original formulas. | Documentation sheets where formulas are displayed alongside calculated values. | Export scenarios that require formulas to be preserved as strings for downstream processing.
// AI Prompts: Generate C# code with Aspose.Cells that writes a formula to A1 and places its FORMULATEXT result in B1, then saves the workbook. | Explain how to enable formula calculation in Aspose.Cells so FORMULATEXT returns the correct string. | Show how to iterate over a range and copy each cell's formula text to an adjacent column using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This C# example creates a new workbook, assigns the formula "=SUM(1,2,3)" to cell A1, uses the Excel FORMULATEXT function in cell B1 to store the literal formula string, forces calculation so B1 displays the text, prints both values to the console, and saves the file as FormulaTextDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Write a formula to cell A1
        cells["A1"].Formula = "=SUM(1,2,3)";

        // Use the Excel FORMULATEXT function to store the formula text of A1 into B1
        cells["B1"].Formula = "=FORMULATEXT(A1)";

        // Calculate all formulas so that B1 shows the formula text
        workbook.CalculateFormula();

        // Display the results
        Console.WriteLine("A1 calculated value: " + cells["A1"].StringValue);
        Console.WriteLine("B1 formula text: " + cells["B1"].StringValue);

        // Save the workbook
        workbook.Save("FormulaTextDemo.xlsx");
    }
}
