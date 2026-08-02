using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsPowerQueryDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSB workbook that may contain Power Query formulas
            string sourcePath = "source.xlsb";

            // Verify that the source file exists before attempting to load it
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                return;
            }

            try
            {
                // Load the workbook (XLSB format is supported by Aspose.Cells)
                Workbook workbook = new Workbook(sourcePath);

                // Use reflection to access Power Query information (avoids compile‑time dependency on specific API versions)
                object mashup = workbook.GetType().GetProperty("DataMashup")?.GetValue(workbook);
                if (mashup != null)
                {
                    // Retrieve the collection of Power Query formulas
                    object formulasObj = mashup.GetType().GetProperty("PowerQueryFormulas")?.GetValue(mashup);
                    if (formulasObj is IEnumerable formulas)
                    {
                        // Count formulas
                        int count = 0;
                        foreach (var _ in formulas) count++;
                        Console.WriteLine($"Number of Power Query formulas: {count}");

                        // Enumerate each formula and display its details
                        foreach (var formula in formulas)
                        {
                            var type = formula.GetType();

                            string name = type.GetProperty("Name")?.GetValue(formula)?.ToString() ?? "N/A";
                            string definition = type.GetProperty("FormulaDefinition")?.GetValue(formula)?.ToString() ?? "N/A";
                            string formulaType = type.GetProperty("Type")?.GetValue(formula)?.ToString() ?? "N/A";
                            string description = type.GetProperty("Description")?.GetValue(formula)?.ToString() ?? "N/A";

                            Console.WriteLine("--------------------------------------------------");
                            Console.WriteLine($"Formula Name       : {name}");
                            Console.WriteLine($"Formula Definition : {definition}");
                            Console.WriteLine($"Formula Type       : {formulaType}");
                            Console.WriteLine($"Description        : {description}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("No Power Query formulas found in the workbook.");
                    }
                }
                else
                {
                    Console.WriteLine("The workbook does not contain Power Query (DataMashup) information.");
                }

                // Optionally, save the workbook (unchanged) to a new file
                string destPath = "output.xlsb";
                workbook.Save(destPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(destPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}