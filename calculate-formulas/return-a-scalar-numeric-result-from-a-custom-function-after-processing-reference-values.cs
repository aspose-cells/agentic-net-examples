// Title: Create a custom REFSUM function in Aspose.Cells for .NET that returns a scalar sum
// Description: Demonstrates building a user‑defined REFSUM function with Aspose.Cells' AbstractCalculationEngine. The sample creates a workbook, writes numbers to A1‑A3, sets the formula =REFSUM(A1:A3) in B1, registers the RefSumEngine, walks through ReferredArea objects, adds the values, and stores the single numeric result in the cell.
// Keywords: Aspose.Cells | C# custom function | AbstractCalculationEngine | REFSUM | scalar result | custom calculation engine | ReferredArea | Excel formula extension | .NET spreadsheet library | user‑defined function
// Common Searches: Aspose.Cells custom REFSUM example | How to add a user defined function in Aspose.Cells C# | Return scalar value from custom calculation engine | Sum range with custom function Aspose.Cells | Implement AbstractCalculationEngine in .NET
// Developer Intent: Add a user‑defined REFSUM function that aggregates a cell range and outputs a single numeric value.
// Use Cases: Financial models that need bespoke aggregation beyond the built‑in SUM. | Domain‑specific calculations such as weighted totals or conditional aggregates. | Cross‑sheet summarization where standard functions are insufficient. | Extending spreadsheet logic in automated reporting pipelines.
// AI Prompts: Write a C# class extending AbstractCalculationEngine to implement REFSUM that sums all numeric cells in a given range. | Show how to configure CalculationOptions with a custom engine and invoke =REFSUM in a worksheet. | Explain error handling for non‑numeric cells in the REFSUM implementation.

using System;
using Aspose.Cells;

// Demonstrates building a user‑defined REFSUM function with Aspose.Cells' AbstractCalculationEngine. The sample creates a workbook, writes numbers to A1‑A3, sets the formula =REFSUM(A1:A3) in B1, registers the RefSumEngine, walks through ReferredArea objects, adds the values, and stores the single numeric result in the cell.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells that will be referenced by the custom function
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(15);

            // Use a custom function REFSUM that sums the values in the given range
            sheet.Cells["B1"].Formula = "=REFSUM(A1:A3)";

            // Configure calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new RefSumEngine()
            };

            // Calculate all formulas in the workbook
            workbook.CalculateFormula(options);

            // Output the scalar numeric result returned by the custom function
            Console.WriteLine("REFSUM result: " + sheet.Cells["B1"].Value);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("RefSumDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Custom calculation engine that implements the REFSUM function
    class RefSumEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Process only the REFSUM function (case‑insensitive)
            if (data.FunctionName.Equals("REFSUM", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;

                // Iterate through all parameters (normally one range)
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Get the parameter value; if it is a range it will be a ReferredArea
                    var area = data.GetParamValue(i) as ReferredArea;

                    if (area != null)
                    {
                        // Sum each cell inside the referred area
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object val = area.GetValue(r, c);
                                if (val != null)
                                {
                                    sum += Convert.ToDouble(val);
                                }
                            }
                        }
                    }
                    else
                    {
                        // Parameter is a plain value (not a range)
                        object val = data.GetParamValue(i);
                        if (val != null)
                        {
                            sum += Convert.ToDouble(val);
                        }
                    }
                }

                // Return the scalar result via CalculatedValue
                data.CalculatedValue = sum;
            }
        }

        // No special handling required for shared formulas
        public override bool ForceRecalculate(string functionName) => false;
    }
}
