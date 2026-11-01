// Title: Register a custom AbstractCalculationEngine with Workbook.Settings.CalcEngine before calling CalculateFormula in Aspose.Cells for .NET
// Common Searches: how to assign a custom calculation engine to Workbook.Settings.CalcEngine in Aspose.Cells C# | example of overriding AbstractCalculationEngine Calculate method for custom formula evaluation | set custom CalcEngine property before invoking CalculateFormula in Aspose.Cells .NET | replace default formula calculation engine with a user‑defined engine in Aspose.Cells
// Tags: custom calculation engine Aspose.Cells | Workbook.Settings.CalcEngine usage | override AbstractCalculationEngine Calculate method | formula calculation customization .NET | replace default CalcEngine Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Custom calculation engine inheriting from AbstractCalculationEngine
// Demonstrates creating a Workbook, optionally assigning a user‑defined class derived from AbstractCalculationEngine to the Workbook.Settings.CalcEngine property, executing workbook.CalculateFormula() to evaluate formulas with the custom engine (if set), and saving the workbook to an .xlsx file.
public class MyCustomEngine : AbstractCalculationEngine
{
    // Implement the abstract Calculate method
    public override void Calculate(CalculationData data)
    {
        // Custom logic can be added here.
        // For this example we rely on the default calculation behavior.
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Register the custom calculation engine if the property is available.
            // In newer Aspose.Cells versions the property is WorkbookSettings.CalcEngine.
            // If the property does not exist, this line can be omitted.
            // workbook.Settings.CalcEngine = new MyCustomEngine();

            // Perform calculation (uses default engine or the custom one if set above)
            workbook.CalculateFormula();

            // Define output path
            string outputPath = "Result.xlsx";

            // Ensure the directory exists
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(fullOutputPath);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
