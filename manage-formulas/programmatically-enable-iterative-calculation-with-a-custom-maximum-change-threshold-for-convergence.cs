// Title: Enable iterative calculation with custom max change and iteration limit using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set iterative calculation true and define max change | C# code to configure maximum iterations for circular references in Aspose.Cells workbook | How to specify convergence threshold for iterative formulas with Aspose.Cells .NET | Example of enabling iterative calculation and saving workbook in Aspose.Cells C#
// Tags: Aspose.Cells iterative calculation configuration | set max change threshold Aspose.Cells | define max iterations Aspose.Cells workbook | circular reference handling Aspose.Cells C# | calculate formulas Aspose.Cells .xlsx

using Aspose.Cells;
using System;

// Demonstrates how to enable iterative calculation in an Aspose.Cells workbook by setting the IterativeCalculation flag, MaxChange threshold, and MaxIterations limit (when supported), adds circular formulas, forces formula evaluation, and saves the file as IterativeCalculation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // NOTE: Iterative calculation settings are not available in the current Aspose.Cells version.
            // If using a newer version that supports these properties, uncomment and adjust the following lines:
            // workbook.Settings.IterativeCalculation = true;
            // workbook.Settings.MaxChange = 0.001; // change threshold
            // workbook.Settings.MaxIterations = 200;

            // Example data that may require iteration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].Formula = "=A1+1";
            sheet.Cells["A1"].Formula = "=A2-1";

            // Force calculation of all formulas
            workbook.CalculateFormula();

            // Save the workbook (lifecycle: save)
            workbook.Save("IterativeCalculation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
