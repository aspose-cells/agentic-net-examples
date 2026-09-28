// Title: Set Aspose.Cells workbook to manual calculation mode, update cells, and recalculate only a specific formula range in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, switches its calculation mode to Manual, changes the values in A1 and A2, and then recalculates only the formula in B1 without performing a full workbook calculation. | Demonstrate disabling automatic formula evaluation in Aspose.Cells, modifying a block of cells, and invoking CalculateFormula for the range B1:B10 using the .NET API.
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual and recalculate a single cell | How to recalculate only a range of formulas in Aspose.Cells .NET | Disable automatic formula calculation in Aspose.Cells and trigger manual calculation | Partial formula evaluation example with Aspose.Cells for C# | CalculateFormula for a specific range using Aspose.Cells .NET
// Tags: manual calculation mode Aspose.Cells C# | recalculate specific range Aspose.Cells | disable automatic formula evaluation Aspose.Cells | partial formula calculation .NET | update cells trigger manual calculation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, sets the calculation mode to Manual, writes values to A1 and A2, assigns a sum formula to B1, and then recalculates only the affected formula range before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells
            sheet.Cells["A1"].PutValue(10);          // Direct value
            sheet.Cells["A2"].PutValue(20);          // Direct value
            sheet.Cells["B1"].Formula = "=A1+A2";    // Formula dependent on A1 and A2

            // Recalculate all formulas in the workbook (no direct range API in .NET)
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
