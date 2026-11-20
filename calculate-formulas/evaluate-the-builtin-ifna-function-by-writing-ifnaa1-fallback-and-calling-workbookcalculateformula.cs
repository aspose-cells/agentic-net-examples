// Title: Using Aspose.Cells for .NET to evaluate the IFNA function with a fallback string in C#
// Common Searches: how to calculate IFNA formula with fallback value using Aspose.Cells C# | Aspose.Cells evaluate Excel IFNA function programmatically | C# example for handling #N/A errors with IFNA in Aspose.Cells | calculate workbook formulas containing NA() using Aspose.Cells .NET | retrieve IFNA result after Workbook.CalculateFormula in C#
// Tags: IFNA formula evaluation Aspose.Cells .NET | handling NA error with IFNA in C# | Workbook.CalculateFormula IFNA example | fallback value for Excel IFNA using Aspose | C# Aspose.Cells error handling in formulas

using Aspose.Cells;
using System;

// The sample creates a new workbook, sets cell A1 to the #N/A error via the NA() function, applies the IFNA(A1, "fallback") formula in cell B1, runs Workbook.CalculateFormula to evaluate all formulas, then reads and prints the fallback string from B1.
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

            // Set A1 to an #N/A error using the NA() function
            sheet.Cells["A1"].Formula = "=NA()";

            // Apply the IFNA formula in B1
            sheet.Cells["B1"].Formula = "=IFNA(A1, \"fallback\")";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the evaluated result from B1
            string result = sheet.Cells["B1"].StringValue;

            // Display the result
            Console.WriteLine($"Result of IFNA: {result}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
