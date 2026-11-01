// Title: Insert a SUM formula into a worksheet cell using Cell.PutValue with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, uses Cell.PutValue to place the formula "=SUM(A1:A10)" into cell B2, and saves the file as Result.xlsx with Aspose.Cells. | Adapt the sample to insert an AVERAGE formula for the range C1:C5 into cell D3 using Cell.PutValue, then export the workbook to an XLSX file. | Show how to apply a custom arithmetic expression such as "=A1*B1+10" to a specific table column cell via Cell.PutValue and save the workbook.
// Common Searches: aspocells c# putvalue formula example | how to set Excel cell formula programmatically using Aspose.Cells | insert SUM(A1:A10) into B2 with Aspose.Cells .NET | using Cell.PutValue to add formulas to table cells in C# | save workbook after adding formula with Aspose.Cells
// Tags: Cell.PutValue set Excel formula | Aspose.Cells insert SUM function | C# add formula to worksheet cell | export workbook after formula insertion | apply custom arithmetic formula Aspose.Cells

using Aspose.Cells;

// // Creates a workbook, places the formula "=SUM(A1:A10)" into cell B2 via Cell.PutValue, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Target cell where the formula will be placed (e.g., B2)
        Cell targetCell = sheet.Cells["B2"];

        // Insert a formula using PutValue with a formula string
        targetCell.PutValue("=SUM(A1:A10)");

        // Save the workbook (using the provided save rule)
        workbook.Save("Result.xlsx");
    }
}
