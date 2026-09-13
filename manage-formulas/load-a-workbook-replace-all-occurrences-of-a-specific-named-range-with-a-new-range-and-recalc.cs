// Title: Replace an existing named range with a new address and recalculate all formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to change the RefersTo property of a specific named range and then call CalculateFormula before saving the workbook. | Programmatically update a named range to a different cell block and trigger a full formula recalculation in a C# Excel automation script.
// Common Searches: Aspose.Cells C# change named range reference and recalc workbook | how to update RefersTo of a named range in Aspose.Cells | recalculate formulas after modifying named range with Aspose.Cells .NET | replace Excel named range programmatically using Aspose.Cells for .NET | C# Aspose.Cells update named range address Sheet1!$C$1:$D$2
// Tags: named range RefersTo update Aspose.Cells | calculate formulas after named range change Aspose.Cells | replace named range address C# | Aspose.Cells workbook modification | Excel named range automation .NET

using Aspose.Cells;
using System;

// Loads input.xlsx, updates the RefersTo property of the named range "OldRange" to "Sheet1!$C$1:$D$2", recalculates all formulas, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Name of the existing named range to be replaced
        string oldRangeName = "OldRange";

        // New range reference (e.g., Sheet1!$C$1:$D$2)
        string newRangeReference = "Sheet1!$C$1:$D$2";

        // Retrieve the named range object
        Name oldName = workbook.Worksheets.Names[oldRangeName];
        if (oldName != null)
        {
            // Update the reference of the named range to the new range
            oldName.RefersTo = newRangeReference;
        }

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
