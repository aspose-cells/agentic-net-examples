using System;
using Aspose.Cells;

public class CompoundInterestEngine : AbstractCalculationEngine
{
    public override void Calculate(CalculationData data)
    {
        // Handle the custom function COMPOUNDINTEREST(principal, rate, periods)
        if (data.FunctionName.Equals("COMPOUNDINTEREST", StringComparison.OrdinalIgnoreCase))
        {
            // Retrieve parameters (they are expected to be numeric)
            double principal = Convert.ToDouble(data.GetParamValue(0));
            double rate = Convert.ToDouble(data.GetParamValue(1));
            double periods = Convert.ToDouble(data.GetParamValue(2));

            // Compound interest formula: principal * (1 + rate) ^ periods
            double result = principal * Math.Pow(1 + rate, periods);

            // Set the calculated value so Aspose.Cells can return it
            data.CalculatedValue = result;
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Input values: principal, rate, number of periods
        cells["A1"].PutValue(1000);   // Principal amount
        cells["B1"].PutValue(0.05);   // Interest rate per period (5%)
        cells["C1"].PutValue(10);     // Number of periods

        // Set the formula that uses the custom function
        cells["D1"].Formula = "=COMPOUNDINTEREST(A1,B1,C1)";

        // Configure calculation options to use the custom engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new CompoundInterestEngine()
        };

        // Perform calculation
        wb.CalculateFormula(opts);

        // Output the result
        Console.WriteLine("Compound Interest Result: " + cells["D1"].Value);
    }
}