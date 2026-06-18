using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ImportArrayElementByIndexDemo
    {
        public static void Run()
        {
            try
            {
                // Load existing workbook if present; otherwise create a new one.
                string sourcePath = "PowerQuerySample.xlsx";
                Workbook workbook = File.Exists(sourcePath) ? new Workbook(sourcePath) : new Workbook();

                // Access Power Query formulas collection.
                var formulas = workbook.DataMashup.PowerQueryFormulas;

                if (formulas.Count == 0)
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                    return;
                }

                // Get first formula and its items.
                var formula = formulas[0];
                var items = formula.PowerQueryFormulaItems;

                // Index of the element to retrieve (zero‑based).
                int targetIndex = 2;
                if (items.Count <= targetIndex)
                {
                    Console.WriteLine($"The formula contains only {items.Count} item(s); index {targetIndex} is out of range.");
                    return;
                }

                var targetItem = items[targetIndex];

                // Display item details.
                Console.WriteLine($"Item at index {targetIndex}:");
                Console.WriteLine($"  Name  : {targetItem.Name}");
                Console.WriteLine($"  Value : {targetItem.Value}");

                // Optionally modify the value.
                // targetItem.Value = "NewValue";

                // Save workbook with any changes.
                workbook.Save("PowerQueryArrayElementDemo.xlsx");
                Console.WriteLine("Workbook saved as PowerQueryArrayElementDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ImportArrayElementByIndexDemo.Run();
        }
    }
}