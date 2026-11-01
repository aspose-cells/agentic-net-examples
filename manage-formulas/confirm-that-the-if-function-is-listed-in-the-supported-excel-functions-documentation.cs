// Title: Check IF function support in Aspose.Cells for .NET by evaluating an IF formula in a workbook
// Common Searches: how to verify that the IF function works in Aspose.Cells using C# | C# code to test Excel IF formula evaluation with Aspose.Cells | determine if IF is listed in Aspose.Cells supported functions documentation | run a workbook calculation to confirm IF function support in Aspose.Cells .NET | sample Aspose.Cells program to check IF formula availability
// Tags: Aspose.Cells IF function support test | C# calculate workbook formulas Aspose.Cells | verify Excel function availability Aspose.Cells | supported functions list Aspose.Cells .NET | evaluate IF formula Aspose.Cells workbook

using System;
using Aspose.Cells;

// The example creates a new Workbook, inserts the formula IF(A1>5, "High", "Low") into cell B1, triggers formula calculation, reads the resulting string, and prints a message confirming that the IF function is listed in Aspose.Cells' Supported Excel Functions documentation when the result is "High".
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a numeric value in cell A1
        sheet.Cells["A1"].PutValue(10);

        // Set a formula using the IF function in cell B1
        // The formula: IF(A1>5, "High", "Low")
        sheet.Cells["B1"].Formula = "IF(A1>5, \"High\", \"Low\")";

        // Calculate the formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the calculated value from B1
        string result = sheet.Cells["B1"].StringValue;

        // Output the result to confirm IF function support
        Console.WriteLine($"Result of IF formula: {result}");

        // If the result is "High", the IF function is supported
        if (result == "High")
        {
            Console.WriteLine("The IF function is listed in the Supported Excel Functions documentation.");
        }
        else
        {
            Console.WriteLine("The IF function may not be supported.");
        }
    }
}
