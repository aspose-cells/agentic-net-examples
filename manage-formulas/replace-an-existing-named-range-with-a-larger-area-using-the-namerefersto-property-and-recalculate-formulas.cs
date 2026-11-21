// Title: Expand an existing named range to a larger area and recalculate all formulas using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change RefersTo of a named range to a larger range | how to expand a named range and recalculate formulas with Aspose.Cells | update named range address in Excel workbook using Aspose.Cells .NET | recalculate all formulas after modifying a named range in Aspose.Cells | replace MyRange with Sheet1!A1:D10 using Aspose.Cells C#
// Tags: modify named range RefersTo Aspose.Cells | expand named range area C# | recalculate workbook formulas Aspose.Cells | update named range address .NET | replace named range with larger range Aspose.Cells

using Aspose.Cells;

// The example loads an Excel workbook, retrieves the named range "MyRange", updates its RefersTo property to a larger cell block (Sheet1!$A$1:$D$10), recalculates all formulas in the workbook, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the named range by its name
        Name namedRange = workbook.Worksheets.Names["MyRange"];
        if (namedRange != null)
        {
            // Replace the named range with a larger area (e.g., A1:D10 on Sheet1)
            namedRange.RefersTo = "Sheet1!$A$1:$D$10";
        }

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
