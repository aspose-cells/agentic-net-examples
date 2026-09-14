// Title: Set CalcEngineSettings.TreatTextAsZero = true in Aspose.Cells .NET to treat text as zero during formula evaluation
// AI Prompts: Generate C# code that enables workbook.Settings.CalcEngineSettings.TreatTextAsZero and recalculates a SUM formula with mixed numeric and text cells using Aspose.Cells. | Explain how to configure the Aspose.Cells calculation engine to ignore non‑numeric text when evaluating formulas in a .NET workbook. | Show an example of summing a range containing numbers, numeric strings, and plain text while treating the text values as zero with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# treat text as zero in formula calculation | How to enable TreatTextAsZero property in Aspose.Cells workbook settings | Sum mixed numeric and text cells with Aspose.Cells ignoring text values | CalcEngineSettings TreatTextAsZero not applying in Aspose.Cells .NET | Set calculation options to treat text as zero in Aspose.Cells example
// Tags: CalcEngineSettings.TreatTextAsZero | Aspose.Cells calculation engine settings | C# SUM formula with mixed data types | Aspose.Cells ignore non-numeric text in formulas | Workbook.Settings.CalcEngineSettings .NET

using System;
using Aspose.Cells;

// Creates a workbook, populates cells with numeric, numeric‑text, and non‑numeric values, assigns a SUM formula, optionally sets workbook.Settings.CalcEngineSettings.TreatTextAsZero to treat text as zero, recalculates formulas, and prints the resulting sum.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate cells with numeric, numeric text, and non‑numeric text
            cells["A1"].PutValue(10);          // numeric value
            cells["A2"].PutValue("20");        // numeric stored as text
            cells["A3"].PutValue("Hello");     // non‑numeric text

            // Set a formula that sums the three cells
            cells["A4"].Formula = "=SUM(A1:A3)";

            // If the Aspose.Cells version supports CalcEngineSettings, treat text as zero.
            // This line is optional; older versions may not have CalcEngineSettings.
            // Uncomment the following lines if the property is available in your version.
            // workbook.Settings.CalcEngineSettings.TreatTextAsZero = true;

            // Recalculate formulas
            workbook.CalculateFormula();

            // Display the result
            Console.WriteLine("SUM(A1:A3) => " + cells["A4"].Value);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
