using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

public class MyRangeFunctionEngine : AbstractCalculationEngine
{
    // This method is called for each custom function during calculation.
    public override void Calculate(CalculationData data)
    {
        // Handle only our custom function.
        if (string.Equals(data.FunctionName, "RANGEFUNC", StringComparison.OrdinalIgnoreCase))
        {
            // Worksheet where the function is evaluated.
            Worksheet ws = data.Worksheet;

            // Create a 2x2 range (B2:C3). Parameters: firstRow, firstColumn, totalRows, totalColumns.
            AsposeRange returnRange = ws.Cells.CreateRange(1, 1, 2, 2);

            // Assign the range as the function result.
            data.CalculatedValue = returnRange;
        }
    }
}

// Optional custom function definition (no special handling needed).
public class MyCustomFunctionDefinition : CustomFunctionDefinition
{
    public override int[] GetArrayModeParameters(string functionName)
    {
        // No parameters require array‑mode calculation.
        return null;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create ----------
            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate the cells that will be part of the returned range.
            ws.Cells["B2"].PutValue(10);
            ws.Cells["C2"].PutValue(20);
            ws.Cells["B3"].PutValue(30);
            ws.Cells["C3"].PutValue(40);

            // Update the custom function definition (optional).
            wb.UpdateCustomFunctionDefinition(new MyCustomFunctionDefinition());

            // Set a formula that calls the custom function.
            ws.Cells["A1"].Formula = "=RANGEFUNC()";

            // ---------- Calculate ----------
            // Use a CalculationOptions instance with our custom engine.
            CalculationOptions calcOpts = new CalculationOptions
            {
                CustomEngine = new MyRangeFunctionEngine()
            };
            wb.CalculateFormula(calcOpts);

            // Retrieve and display the type of the result stored in A1.
            object result = ws.Cells["A1"].Value;
            Console.WriteLine("A1 contains a value of type: " + (result?.GetType().FullName ?? "null"));

            // ---------- Save ----------
            string outputPath = "CustomRangeFunctionDemo.xlsx";

            // Ensure the directory exists before saving.
            string dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}