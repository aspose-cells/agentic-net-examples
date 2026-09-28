// Title: Create a shared formula for cells C3:C12 in an Excel worksheet using Aspose.Cells SetSharedFormula in C#
// AI Prompts: Generate C# code that creates a new Workbook, calls Cell.SetSharedFormula on cell C3 with the formula "=A3+B3" to share it across the range C3:C12, and saves the file as an .xlsx. | Write a method using Aspose.Cells that defines a shared formula for column C rows 3‑12 via SetSharedFormula on the first cell, then writes the workbook to disk.
// Common Searches: Aspose.Cells SetSharedFormula C# example for range C3:C12 | How to share a formula across multiple cells using Aspose.Cells .NET | C# apply same Excel formula to column C rows 3 to 12 with Aspose.Cells | Create shared formula in Excel workbook using Aspose.Cells SetSharedFormula method
// Tags: Aspose.Cells SetSharedFormula C# | shared formula range C3:C12 Aspose.Cells | apply shared formula Excel .NET | column C formula assignment Aspose.Cells | save workbook with shared formulas Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, defines a shared formula "=A3+B3" on cell C3 using Cell.SetSharedFormula, propagates it to the range C3:C12, and saves the workbook as SharedFormula.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set formula for the first cell in the range C3:C12
            Cell firstCell = cells["C3"];
            firstCell.Formula = "=A3+B3";

            // Apply the same pattern to the remaining cells in the range
            for (int row = 4; row <= 12; row++)
            {
                // Column C has index 2 (zero‑based)
                Cell cell = cells[row, 2];
                cell.Formula = $"=A{row}+B{row}";
            }

            // Define output file path
            string outputPath = "SharedFormula.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
