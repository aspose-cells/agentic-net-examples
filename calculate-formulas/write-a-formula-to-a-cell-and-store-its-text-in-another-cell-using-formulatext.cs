// Title: Write a SUM formula to a cell and copy its formula text to another cell using Aspose.Cells for .NET (C#)
// Common Searches: how to assign a formula to a cell and retrieve its text with Aspose.Cells in C# | Aspose.Cells C# write SUM formula to A1 and copy formula string to B1 | retrieve formula text from a worksheet cell using Aspose.Cells .NET API | store a cell's formula as plain text in a different cell with Aspose.Cells | example of reading and writing formula text in an Excel file using Aspose.Cells for C#
// Tags: write formula to cell Aspose.Cells C# | read formula text Aspose.Cells | copy formula string between cells Aspose.Cells | Aspose.Cells SUM formula example | save workbook output.xlsx Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, writes the formula "=SUM(B1:B5)" into cell A1, reads the formula text via the Formula property, places that text into cell B1, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Write a formula to cell A1
        Cell formulaCell = sheet.Cells["A1"];
        formulaCell.Formula = "=SUM(B1:B5)";

        // Retrieve the formula text using the Formula property (acts as FormulaText)
        string formulaText = formulaCell.Formula;

        // Store the formula text in another cell (e.g., B1)
        Cell textCell = sheet.Cells["B1"];
        textCell.PutValue(formulaText);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("output.xlsx");
    }
}
