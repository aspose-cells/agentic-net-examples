// Title: Read the formula in cell E3 of the second worksheet after deleting a row with Aspose.Cells for .NET
// AI Prompts: Delete the first row of the second worksheet, let Aspose.Cells shift the cells, and then fetch the updated formula from cell E3 in C#. | Provide C# code that removes a row, updates all references, and returns the formula contained in E3 on the second sheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# read cell formula after row deletion | How to get updated formula in Excel when a row is removed using Aspose.Cells | Retrieve E3 formula after deleting first row in second worksheet Aspose.Cells .NET | C# Aspose.Cells preserve formulas while deleting rows | Read formula of a shifted cell after row removal with Aspose.Cells
// Tags: delete row update formulas Aspose.Cells | read cell formula Aspose.Cells C# | second worksheet cell E3 retrieval | row deletion reference shift Aspose.Cells | C# workbook modify rows preserve formulas

using System;
using Aspose.Cells;

// The program loads an Excel file, accesses the second worksheet, deletes the first row while shifting cells, reads the formula from cell E3 after the deletion, prints it, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];

        // Delete a row (example: first row) and update references
        sheet.Cells.DeleteRow(0, true);

        // Read the formula from cell E3 after the deletion
        string formula = sheet.Cells["E3"].Formula;

        // Output the formula
        Console.WriteLine($"Formula in E3 after deletion: {formula}");

        // Save the workbook if needed
        workbook.Save("output.xlsx");
    }
}
