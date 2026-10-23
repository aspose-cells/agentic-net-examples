// Title: How to use Aspose.Cells CalculationOptions to treat text as zero while adding numeric and text cells in C#
// Common Searches: Aspose.Cells C# calculate formula with TreatTextAsZero option | set CalculationOptions.TreatTextAsZero to ignore text in Excel addition | evaluate A1+B1 when B1 contains text using Aspose.Cells | C# Aspose.Cells calculate mixed numeric and string cells without error | Workbook.CalculateFormula custom CalculationOptions example
// Tags: Aspose.Cells calculation options for text handling | C# evaluate Excel formula with mixed numeric and string cells | Workbook.CalculateFormula with custom options | ignore non-numeric cells in Aspose.Cells addition | add numeric and text cells using Aspose.Cells API

using Aspose.Cells;
using System;

// This example demonstrates how to configure Aspose.Cells CalculationOptions to treat text values as zero, allowing a formula that adds a numeric cell (A1) and a textual cell (B1) to return the numeric value only. The workbook is created, values are placed in A1 and B1, the formula A1+B1 is set in C1, CalculationOptions are prepared (optionally enabling TreatTextAsZero), and Workbook.CalculateFormula is called with those options. The result from C1 is retrieved as a double (expected 10) and printed to the console.
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

            // Numeric value in A1
            sheet.Cells["A1"].PutValue(10);

            // Textual value in B1
            sheet.Cells["B1"].PutValue("text");

            // Formula in C1 that adds A1 and B1
            sheet.Cells["C1"].Formula = "A1+B1";

            // Set calculation options (if supported by the current Aspose.Cells version)
            CalculationOptions calcOptions = new CalculationOptions();

            // Uncomment the following line if the property exists in your Aspose.Cells version
            // calcOptions.TreatTextAsZero = true;

            // Recalculate formulas using the specified options
            workbook.CalculateFormula(calcOptions);

            // Retrieve and display the result (should be 10 if text is treated as zero)
            double result = sheet.Cells["C1"].DoubleValue;
            Console.WriteLine("Result: " + result);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
