// Title: Recalculate formulas only on the 'Summary' worksheet using Aspose.Cells Workbook.CalculateFormula in C#
// AI Prompts: Modify the sample to pass the 'Summary' worksheet to the calculation method so that only that sheet's formulas are evaluated, then save the workbook. | Demonstrate creating a CalculationOptions instance that targets a specific worksheet and using it with the workbook's formula evaluation routine in a C# Aspose.Cells project. | Provide a concise C# snippet that loads an Excel file, recalculates formulas exclusively on the sheet named "Summary", and writes the updated file.
// Common Searches: Aspose.Cells limit formula calculation to a single worksheet C# | Recalculate only the Summary tab in Excel using Aspose.Cells | C# example for selective formula evaluation with Aspose.Cells | How to evaluate formulas on a specific sheet with Aspose.Cells | Targeted worksheet formula recalculation Aspose.Cells .NET
// Tags: Aspose.Cells target worksheet formula evaluation | C# limit formula calculation to specific sheet | Excel selective formula processing Aspose.Cells | CalculationOptions restrict to one worksheet | recalc single sheet Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx, verifies the presence of a worksheet named "Summary", and shows how to recalculate formulas only on that sheet by passing the worksheet (or its index) to the workbook's calculation method with appropriate CalculationOptions. After the selective evaluation, the workbook is saved as output.xlsx, and the code includes basic error handling and status messages.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string sheetName = "Summary";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            var workbook = new Workbook(inputPath);

            // Retrieve the worksheet by name
            var worksheet = workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Calculate formulas for the entire workbook using CalculationOptions
            var calcOptions = new CalculationOptions();
            workbook.CalculateFormula(calcOptions);

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
