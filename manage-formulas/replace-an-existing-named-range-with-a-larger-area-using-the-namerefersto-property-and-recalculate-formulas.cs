// Title: Expand a named range to a larger area and recalculate all formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Set the Name.RefersTo property to a new cell range and invoke Workbook.CalculateFormula to refresh every formula in a C# Aspose.Cells project. | Replace an existing named range with a bigger range definition and trigger a full workbook recalculation using the Aspose.Cells API.
// Common Searches: how to change the RefersTo value of a named range in Aspose.Cells C# | expand an existing named range to a larger cell block and recalc formulas with Aspose.Cells | C# Aspose.Cells update named range area and recalculate workbook | Aspose.Cells replace MyRange with Sheet1!A1:D10 and recalculate formulas | programmatically increase named range size in Excel using Aspose.Cells .NET
// Tags: Aspose.Cells update Name.RefersTo property | C# expand named range Excel | Aspose.Cells recalculate workbook formulas | modify named range area programmatically | Aspose.Cells set named range to larger range

using Aspose.Cells;

// Loads input.xlsx, expands the named range 'MyRange' to Sheet1!A1:D10 via Name.RefersTo, recalculates all formulas, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the named range by its name
        Name namedRange = workbook.Worksheets.Names["MyRange"];
        if (namedRange != null)
        {
            // Replace the existing range with a larger area using RefersTo property
            // Example: expand to A1:D10 on Sheet1
            namedRange.RefersTo = "='Sheet1'!$A$1:$D$10";
        }

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the modified workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
