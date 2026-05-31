using System;
using Aspose.Cells;

public class VerifyCustomEngineReset
{
    // Simple custom calculation engine that implements a user‑defined function MYFUNC
    private class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function MYFUNC
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Retrieve the two parameters passed to MYFUNC
                double param1 = Convert.ToDouble(data.GetParamValue(0));
                double param2 = Convert.ToDouble(data.GetParamValue(1));

                // Return the sum of the parameters as the calculated value
                data.CalculatedValue = param1 + param2;
            }
        }
    }

    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new workbook and set up test data
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells A1 and A2 with numeric values
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);

        // Use the custom function MYFUNC in cell A3
        sheet.Cells["A3"].Formula = "=MYFUNC(A1, A2)";

        // -------------------------------------------------
        // 2. Assign a custom calculation engine and calculate
        // -------------------------------------------------
        CalculationOptions optionsWithEngine = new CalculationOptions
        {
            CustomEngine = new MyCustomEngine()   // Attach custom engine
        };

        workbook.CalculateFormula(optionsWithEngine);
        Console.WriteLine("Result with custom engine: " + sheet.Cells["A3"].Value); // Expected 15

        // -------------------------------------------------
        // 3. Reset calculation options to default (no custom engine)
        // -------------------------------------------------
        CalculationOptions defaultOptions = new CalculationOptions(); // CustomEngine is null by default

        // Re‑calculate to ensure the workbook works without the custom engine
        workbook.CalculateFormula(defaultOptions);

        // Verify that the CustomEngine property is null after reset
        Console.WriteLine("CustomEngine after reset is null: " + (defaultOptions.CustomEngine == null));

        // -------------------------------------------------
        // 4. Save the workbook (lifecycle rule: save)
        // -------------------------------------------------
        string filePath = "VerifyCustomEngine.xlsx";
        workbook.Save(filePath);

        // -------------------------------------------------
        // 5. Load the workbook back (lifecycle rule: load)
        // -------------------------------------------------
        Workbook loadedWorkbook = new Workbook(filePath);
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // Use default calculation options again
        CalculationOptions loadedOptions = new CalculationOptions();
        loadedWorkbook.CalculateFormula(loadedOptions);

        // Verify that after loading, no custom engine is attached
        Console.WriteLine("After load, CustomEngine is null: " + (loadedOptions.CustomEngine == null));
        Console.WriteLine("Result after load (should be 15): " + loadedSheet.Cells["A3"].Value);
    }
}