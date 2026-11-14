// Title: Create a conditional warning in Excel with Aspose.Cells for .NET using the IF function for minimum value validation
// Common Searches: Aspose.Cells C# set IF formula to display warning when cell value is less than a threshold | How to programmatically add a conditional warning in an Excel workbook using Aspose.Cells .NET | Validate numeric cell against a minimum value with IF function in Aspose.Cells for C# | Calculate workbook formulas after inserting an IF condition with Aspose.Cells in .NET | Generate warning message in Excel using Aspose.Cells when a value falls below a defined minimum
// Tags: programmatic IF condition Aspose.Cells .NET | minimum value check Excel Aspose.Cells | numeric threshold validation Aspose.Cells | evaluate workbook formulas Aspose.Cells | save workbook as XLSX Aspose.Cells

// Include Aspose.Cells namespace
using Aspose.Cells;

// The example creates a new workbook, writes a numeric value (5.0) to cell B2, defines a minimum threshold of 10.0, and assigns an IF formula to cell C2 that displays a warning message when B2 is below the threshold. It then evaluates all formulas and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Define the minimum acceptable value
        double minValue = 10.0;

        // Example numeric field placed in cell B2
        sheet.Cells["B2"].PutValue(5.0); // This value is below the minimum

        // Use the IF function to display a warning when the value is below the minimum
        // The result will appear in cell C2
        sheet.Cells["C2"].Formula = $"IF(B2<{minValue},\"Warning: Value below minimum\",\"\")";

        // Evaluate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
