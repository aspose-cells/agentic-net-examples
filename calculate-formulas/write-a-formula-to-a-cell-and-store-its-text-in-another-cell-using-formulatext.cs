// Title: Write a formula to a cell and copy its formula text to another cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# write SUM formula to A1 and get formula string | How to retrieve formula text from a cell with Aspose.Cells .NET | Copy formula expression to another cell using Aspose.Cells API | Store cell formula as plain text in Excel file with Aspose.Cells C# | Example of using Formula property to read and write formula text in Aspose.Cells
// Tags: Aspose.Cells write formula C# | Aspose.Cells retrieve formula text | copy formula expression between cells | Formula property for Excel manipulation | save workbook with Aspose.Cells after formula

using Aspose.Cells;

// Creates a new workbook, writes the formula "=SUM(B1:B5)" to cell A1, reads the formula text via the Formula property, places that text into cell B1, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Write a formula to cell A1
        Cell formulaCell = cells["A1"];
        formulaCell.Formula = "=SUM(B1:B5)";

        // Retrieve the formula text using the Formula property (FormulaText)
        string formulaText = formulaCell.Formula;

        // Store the formula text in cell B1
        Cell textCell = cells["B1"];
        textCell.PutValue(formulaText);

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
