// Title: Enable printing of gridlines only on the first three worksheets in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable gridline printing on specific worksheets | How to turn on print gridlines for the first three sheets of an Excel file with Aspose.Cells | C# set PrintGridlines property for selected worksheets using Aspose.Cells | Selective worksheet printing settings Aspose.Cells .NET | Limit gridline printing to certain worksheets in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells selective PrintGridlines setting | C# worksheet PageSetup.PrintGridlines configuration | Excel workbook modify printing options Aspose.Cells | per‑sheet gridline printing Aspose.Cells .NET | adjust worksheet print settings programmatically

using System;
using Aspose.Cells;

// The example loads an existing workbook, iterates over up to three worksheets, sets each worksheet's PageSetup.PrintGridlines flag to true, and saves the workbook, thereby enabling gridline printing only on those selected sheets while leaving other sheets unchanged.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Determine how many worksheets to modify (up to the first three)
        int sheetsToProcess = Math.Min(3, workbook.Worksheets.Count);

        // Enable printing of gridlines for the first three worksheets
        for (int i = 0; i < sheetsToProcess; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            // Set the PrintGridlines flag to true for this worksheet
            sheet.PageSetup.PrintGridlines = true;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
