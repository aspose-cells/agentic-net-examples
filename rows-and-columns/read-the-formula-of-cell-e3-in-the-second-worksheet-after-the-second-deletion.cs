// Title: Retrieve the formula of cell E3 from the second worksheet after deleting the first two rows using Aspose.Cells for .NET
// AI Prompts: Delete the first two rows of the second worksheet and return the formula stored in cell E3 with Aspose.Cells in C#. | Provide C# code that removes the top two rows of the second sheet and then reads the formula from cell E3.
// Common Searches: Aspose.Cells C# read formula from cell after deleting rows in second worksheet | how to get E3 formula after removing first rows using Aspose.Cells .NET | C# Aspose.Cells delete top rows then retrieve cell formula example
// Tags: Aspose.Cells delete rows second worksheet | Aspose.Cells read cell formula C# | retrieve formula after row deletion .NET | cell E3 formula extraction Aspose.Cells

using Aspose.Cells;

// Loads a workbook, accesses the second worksheet, deletes the first two rows, reads the formula from cell E3, and writes it to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];

        // First deletion (e.g., delete the first row)
        sheet.Cells.DeleteRows(0, 1);

        // Second deletion (delete the new first row after the previous deletion)
        sheet.Cells.DeleteRows(0, 1);

        // Read the formula from cell E3 (row 2, column 4)
        string formula = sheet.Cells["E3"].Formula;

        // Output the formula
        System.Console.WriteLine(formula);
    }
}
