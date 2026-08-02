using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryUpdate
{
    public class UpdatePowerQuerySource
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains Power Query formulas
            Workbook workbook = new Workbook(inputPath);

            // Define the new CSV file path you want the Power Query to reference
            string newCsvPath = @"C:\Data\NewSourceFile.csv";

            // Verify the new CSV source exists
            if (!File.Exists(newCsvPath))
            {
                Console.WriteLine($"CSV source file not found: {newCsvPath}");
                return;
            }

            // Iterate through all Power Query formulas in the workbook
            foreach (PowerQueryFormula formula in workbook.DataMashup.PowerQueryFormulas)
            {
                // Iterate through the items of each formula
                foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                {
                    // Typically the source file path is stored in an item named "Source"
                    if (item.Name.Equals("Source", StringComparison.OrdinalIgnoreCase))
                    {
                        // Locate the first quoted string (the file path) and replace it
                        int firstQuote = item.Value.IndexOf('\"');
                        int secondQuote = item.Value.IndexOf('\"', firstQuote + 1);
                        if (firstQuote >= 0 && secondQuote > firstQuote)
                        {
                            string before = item.Value.Substring(0, firstQuote + 1);
                            string after = item.Value.Substring(secondQuote);
                            item.Value = before + newCsvPath + after;
                        }
                        else
                        {
                            // Unexpected format – replace the whole value with the new path
                            item.Value = newCsvPath;
                        }

                        Console.WriteLine($"Updated Power Query item '{item.Name}' to new path: {newCsvPath}");
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
    }
}