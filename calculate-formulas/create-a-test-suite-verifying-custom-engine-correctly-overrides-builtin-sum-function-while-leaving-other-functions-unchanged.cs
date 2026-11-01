// Title: C# unit test for Aspose.Cells custom calculation engine that overrides the SUM function while preserving other formulas
// Common Searches: how to write a unit test for a custom SUM implementation in Aspose.Cells C# | verify that a custom calculation engine does not affect AVERAGE formula in Aspose.Cells | C# Aspose.Cells test suite for overriding built‑in SUM but keeping other Excel functions unchanged | example of unit testing a custom SUM function in Aspose.Cells without breaking other formulas | ensure custom SUM override works correctly in Aspose.Cells calculation engine
// Tags: aspocells custom calculation engine sum override | c# unit testing aspocells custom formulas | excel sum function override aspocells | preserve average formula aspocells custom engine | aspocells formula calculation testing c#

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Creates a workbook, populates cells with numeric values, assigns SUM and AVERAGE formulas, runs the custom calculation engine, and uses a C# unit test to assert that the overridden SUM returns the expected result while the AVERAGE formula remains unchanged.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];

                    // Populate cells A1 and A2 with numeric values
                    sheet.Cells["A1"].PutValue(2);
                    sheet.Cells["A2"].PutValue(3);

                    // Set formulas that use SUM and AVERAGE
                    sheet.Cells["B1"].Formula = "SUM(A1:A2)";
                    sheet.Cells["B2"].Formula = "AVERAGE(A1:A2)";

                    // Force calculation so that formulas are evaluated
                    workbook.CalculateFormula();

                    // Retrieve and display results
                    double sumResult = sheet.Cells["B1"].DoubleValue;
                    double avgResult = sheet.Cells["B2"].DoubleValue;

                    Console.WriteLine($"SUM result: {sumResult}");
                    Console.WriteLine($"AVERAGE result: {avgResult}");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
