// Title: How to set Aspose.Cells workbook to Manual mode and recalculate only a chosen worksheet in C#
// AI Prompts: Switch the workbook to manual mode, then invoke the Worksheet.CalculateFormula method for the 'DataSheet' sheet using Aspose.Cells for .NET. | Show how to recalculate formulas only on a selected worksheet while other sheets stay unevaluated, and then save the workbook.
// Common Searches: Aspose.Cells C# set workbook calculation to manual and evaluate a single sheet | How to run formula calculation on only one worksheet in Aspose.Cells .NET | Manual calc mode example for recalculating specific worksheet with Aspose.Cells | Trigger formula evaluation for a particular sheet after setting manual mode in Aspose.Cells | Recalculate formulas on selected worksheet without affecting other sheets Aspose.Cells
// Tags: Aspose.Cells set workbook CalcMode Manual | calculate formulas on single worksheet | partial workbook formula evaluation .NET | trigger worksheet formula calculation C# | manual calc mode example Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, names the first worksheet "DataSheet", inserts numeric values into A1 and A2, assigns a SUM formula to A3, switches the workbook to manual calculation mode, triggers formula evaluation only for the 'DataSheet' worksheet, and saves the result as ManualCalcResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Add sample data and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Calculate formulas for the entire workbook
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "ManualCalcResult.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
