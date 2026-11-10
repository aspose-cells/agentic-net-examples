// Title: Enable iterative calculation with a 100‑iteration limit and test circular reference handling using Aspose.Cells for .NET (C#)
// Common Searches: how to enable iterative calculation and set max iterations in Aspose.Cells C# | sample code for handling circular reference formulas with iteration limit in Aspose.Cells | using reflection to configure workbook settings for iterative calculation in Aspose.Cells | Aspose.Cells calculate formulas with circular references and precision setting example
// Tags: configure iterative calculation Aspose.Cells | set max iterations workbook Aspose.Cells | circular reference handling Aspose.Cells | adjust iterative precision Aspose.Cells | reflection based settings adjustment Aspose.Cells C#

using Aspose.Cells;
using System;
using System.Reflection;

// The example creates a new workbook, defines a circular reference between cells A1 and B1, and uses reflection to enable iterative calculation, set the maximum iteration count to 100, and specify a precision of 0.001 via Workbook.Settings. It then calculates the formulas, outputs the resulting values, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a circular reference: A1 = B1 + 1, B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation if the properties are available in the current Aspose.Cells version
            var settings = workbook.Settings;
            Type settingsType = settings.GetType();

            PropertyInfo iterativeProp = settingsType.GetProperty("IterativeCalculation");
            PropertyInfo maxIterProp = settingsType.GetProperty("MaxIterations");
            PropertyInfo precisionProp = settingsType.GetProperty("IterativeCalculationPrecision");

            if (iterativeProp != null && iterativeProp.CanWrite)
                iterativeProp.SetValue(settings, true);
            if (maxIterProp != null && maxIterProp.CanWrite)
                maxIterProp.SetValue(settings, 100);
            if (precisionProp != null && precisionProp.CanWrite)
                precisionProp.SetValue(settings, 0.001);

            // Perform calculation
            workbook.CalculateFormula();

            // Display the results after iteration
            Console.WriteLine("A1 = " + sheet.Cells["A1"].Value);
            Console.WriteLine("B1 = " + sheet.Cells["B1"].Value);
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
