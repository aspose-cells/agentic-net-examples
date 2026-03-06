using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MaxRowsOfSharedFormulaScenarios
    {
        public static void Run()
        {
            // Load an existing XLSX workbook with default formula parsing
            string inputPath = "input.xlsx";
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = true // parse formulas while loading
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Show the default maximum rows allowed for a shared formula
            Console.WriteLine("Default MaxRowsOfSharedFormula: " + workbook.Settings.MaxRowsOfSharedFormula);

            // -------------------------------------------------
            // Scenario 1: Restrict shared formula rows to a small number
            // -------------------------------------------------
            workbook.Settings.MaxRowsOfSharedFormula = 50; // limit to 50 rows
            Worksheet sheet1 = workbook.Worksheets[0];
            Cells cells1 = sheet1.Cells;

            // Attempt to set a shared formula that exceeds the limit (60 rows)
            try
            {
                cells1["B1"].SetSharedFormula("=A1", 60, 1);
                Console.WriteLine("Shared formula set for 60 rows with limit 50.");
                Console.WriteLine("Formula in B60: " + cells1["B60"].Formula);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error when setting shared formula exceeding limit: " + ex.Message);
            }

            // -------------------------------------------------
            // Scenario 2: Increase the limit to accommodate a larger range
            // -------------------------------------------------
            workbook.Settings.MaxRowsOfSharedFormula = 200; // raise limit
            Worksheet sheet2 = workbook.Worksheets.Add("LargeFormulaSheet");
            Cells cells2 = sheet2.Cells;

            // Set a shared formula for 150 rows (within the new limit)
            cells2["C1"].SetSharedFormula("=A1*2", 150, 1);
            Console.WriteLine("Shared formula set for 150 rows after increasing limit.");
            Console.WriteLine("Formula in C150: " + cells2["C150"].Formula);

            // -------------------------------------------------
            // Scenario 3: Improve load performance for very large sheets
            // -------------------------------------------------
            LoadOptions fastLoadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false // skip formula parsing on load
            };
            Workbook fastWorkbook = new Workbook(inputPath, fastLoadOptions);
            fastWorkbook.Settings.MaxRowsOfSharedFormula = 5000; // set a reasonable limit for this workbook
            Console.WriteLine("Loaded workbook with ParsingFormulaOnOpen = false. MaxRowsOfSharedFormula set to " + fastWorkbook.Settings.MaxRowsOfSharedFormula);

            // -------------------------------------------------
            // Save the modified workbook
            // -------------------------------------------------
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Modified workbook saved to " + outputPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MaxRowsOfSharedFormulaScenarios.Run();
        }
    }
}