using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ModifyPowerQueryFormula
    {
        public static void Run()
        {
            // Determine workbook path relative to the executable directory
            string workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "default.xlsx");

            // If the workbook does not exist, create an empty one
            if (!File.Exists(workbookPath))
            {
                var emptyWb = new Workbook();
                emptyWb.Save(workbookPath);
            }

            // Load the workbook that may contain Power Query formulas
            Workbook workbook = new Workbook(workbookPath);

            // Get DataMashup object (may be null if workbook has no Power Query)
            var dataMashup = workbook.DataMashup;
            if (dataMashup == null || dataMashup.PowerQueryFormulas == null || dataMashup.PowerQueryFormulas.Count == 0)
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
                return;
            }

            // Get the first Power Query formula
            var formula = dataMashup.PowerQueryFormulas[0];
            if (formula == null || formula.PowerQueryFormulaItems == null || formula.PowerQueryFormulaItems.Count == 0)
            {
                Console.WriteLine("The selected Power Query formula has no items.");
                return;
            }

            // Access the first item of the formula
            var item = formula.PowerQueryFormulaItems[0];

            // Modify the item's value (example: replace a drive letter in a path)
            string originalValue = item.Value;
            string modifiedValue = originalValue.Replace(@"C:\", @"D:\");
            item.Value = modifiedValue;

            // Output the change for verification
            Console.WriteLine($"Item Name: {item.Name}");
            Console.WriteLine($"Original Value: {originalValue}");
            Console.WriteLine($"Modified Value: {modifiedValue}");

            // Save the workbook with the updated Power Query formula item
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "default_modified.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ModifyPowerQueryFormula.Run();
        }
    }
}