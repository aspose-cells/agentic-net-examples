// Title: Get the formula from cell E3 on the second worksheet with Aspose.Cells for .NET before removing the sheet
// AI Prompts: Use Aspose.Cells in C# to read the Formula property of cell E3 on the worksheet at index 1, output the string, then delete that worksheet and save the workbook. | Retrieve the Excel formula stored in E3 of the second sheet, capture it before calling Worksheets.RemoveAt, and write the result to the console.
// Common Searches: aspnet aspose.cells read cell formula before deleting worksheet | c# get formula from E3 on second sheet then remove sheet | how to extract formula from a specific cell using Aspose.Cells before worksheet removal | retrieve Excel cell formula with Aspose.Cells .NET prior to worksheet deletion
// Tags: aspocells read cell formula c# | aspocells worksheet index access | aspocells delete worksheet after read | excel formula extraction aspocells .net | c# get cell formula before worksheet removal

using Aspose.Cells;

// The example loads an Excel workbook, accesses the second worksheet (index 1), reads the formula from cell E3, prints it, removes the second worksheet, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];

        // Get cell E3
        Cell cell = sheet.Cells["E3"];

        // Retrieve the formula before any deletion
        string formula = cell.Formula;

        // Display the formula (optional, for verification)
        System.Console.WriteLine("Formula in E3 before deletion: " + formula);

        // Delete the second worksheet
        workbook.Worksheets.RemoveAt(1);

        // Save the workbook after deletion (save rule)
        workbook.Save("output.xlsx");
    }
}
