using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaTracingDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "SampleWithFormulas.xlsx";

            // Ensure the sample file exists; if not, create a simple workbook with a formula.
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempWs = tempWb.Worksheets[0];
                tempWs.Cells["A1"].Formula = "=SUM(1,2)";
                tempWb.Save(sourcePath);
            }

            // ------------------------------------------------------------
            // 1. Load the workbook with formula parsing enabled (default)
            // ------------------------------------------------------------
            LoadOptions loadWithParsing = new LoadOptions();
            loadWithParsing.ParsingFormulaOnOpen = true; // Parse formulas while loading
            Workbook wbParsed = new Workbook(sourcePath, loadWithParsing);

            Console.WriteLine("=== Workbook loaded with ParsingFormulaOnOpen = true ===");
            DisplayFormulaInfo(wbParsed);

            // ------------------------------------------------------------
            // 2. Load the same workbook with formula parsing disabled
            // ------------------------------------------------------------
            LoadOptions loadWithoutParsing = new LoadOptions();
            loadWithoutParsing.ParsingFormulaOnOpen = false; // Skip parsing formulas on load
            Workbook wbUnparsed = new Workbook(sourcePath, loadWithoutParsing);

            Console.WriteLine("\n=== Workbook loaded with ParsingFormulaOnOpen = false ===");
            // At this point formulas are stored as raw strings and not parsed.
            // Accessing the Formula property will return the original string,
            // while the Value property will be default (e.g., 0 or null).
            DisplayFormulaInfo(wbUnparsed);

            // ------------------------------------------------------------
            // 3. Manually parse the unparsed formulas after loading
            // ------------------------------------------------------------
            // The ParseFormulas method parses all formulas that were not parsed
            // during the load operation. The boolean argument indicates whether
            // to ignore errors (true) or throw an exception on the first invalid formula (false).
            wbUnparsed.ParseFormulas(ignoreError: true);
            Console.WriteLine("\n=== After calling ParseFormulas on the unparsed workbook ===");
            DisplayFormulaInfo(wbUnparsed);

            // ------------------------------------------------------------
            // 4. Inspect and modify FormulaSettings (calculation related)
            // ------------------------------------------------------------
            FormulaSettings settings = wbParsed.Settings.FormulaSettings;

            Console.WriteLine("\n=== FormulaSettings before modification ===");
            Console.WriteLine($"CalculateOnOpen          : {settings.CalculateOnOpen}");
            Console.WriteLine($"EnableCalculationChain   : {settings.EnableCalculationChain}");
            Console.WriteLine($"CalculationMode          : {settings.CalculationMode}");

            // Change some settings to demonstrate their effect
            settings.CalculateOnOpen = true;                     // Request full calculation on open (saved to file)
            settings.EnableCalculationChain = true;              // Enable calculation chain for faster repeated calculations
            settings.CalculationMode = CalcModeType.Manual;      // Switch to manual mode

            Console.WriteLine("\n=== FormulaSettings after modification ===");
            Console.WriteLine($"CalculateOnOpen          : {settings.CalculateOnOpen}");
            Console.WriteLine($"EnableCalculationChain   : {settings.EnableCalculationChain}");
            Console.WriteLine($"CalculationMode          : {settings.CalculationMode}");

            // ------------------------------------------------------------
            // 5. Save the workbook to observe that the modified settings are persisted
            // ------------------------------------------------------------
            string outputPath = "ModifiedFormulaSettings.xlsx";
            wbParsed.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved with updated FormulaSettings to '{outputPath}'.");
        }

        // Helper method to display formula and value of a sample cell (A1) in the first worksheet
        static void DisplayFormulaInfo(Workbook workbook)
        {
            Worksheet ws = workbook.Worksheets[0];
            Cell cell = ws.Cells["A1"];

            // Formula property shows the stored formula string (parsed or raw)
            Console.WriteLine($"Cell A1 Formula : {cell.Formula}");

            // Value property shows the evaluated result if the formula has been parsed and calculated
            // If parsing was skipped and no calculation performed, this may be default (e.g., 0)
            Console.WriteLine($"Cell A1 Value   : {cell.Value}");
        }
    }
}