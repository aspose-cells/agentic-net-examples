// Title: C# – Update Power Query CSV source path in an Excel workbook with Aspose.Cells
// Description: Loads an Excel file, checks for a DataMashup containing Power Query formulas, iterates each PowerQueryFormulaItem, replaces the quoted CSV file path in the M expression with a new location using a regular expression, and saves the workbook as a new file. Includes robust error handling for missing files and save failures.
// Keywords: Aspose.Cells Power Query update | C# replace CSV path in PowerQueryFormulaItem | DataMashup edit Excel | modify Power Query M expression | .NET change Power Query source file | Windows Excel automation Aspose | batch update Power Query connections
// Common Searches: how to change csv file path in power query using Aspose.Cells | c# update PowerQueryFormulaItem value | replace file path in Power Query M expression programmatically | Aspose.Cells edit DataMashup Power Query formulas | update Excel Power Query source file path .NET
// Developer Intent: Programmatically replace the existing CSV reference in every Power Query formula of an Excel workbook with a new file path and persist the changes.
// Use Cases: Migrate legacy reports to a new data folder without manual editing. | Automate source‑file updates across dozens of workbooks before distribution. | Integrate path‑rewriting into CI/CD pipelines for Excel‑based reporting solutions.
// AI Prompts: Write C# code that uses Aspose.Cells to locate all PowerQueryFormulaItem objects and substitute the quoted CSV path with a user‑provided string. | Explain safe practices for editing Power Query M expressions via the DataMashup API, ensuring other formulas remain untouched. | Provide comprehensive error‑handling patterns for loading workbooks, iterating PowerQueryFormulas, applying regex replacements, and saving the updated file.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Loads an Excel file, checks for a DataMashup containing Power Query formulas, iterates each PowerQueryFormulaItem, replaces the quoted CSV file path in the M expression with a new location using a regular expression, and saves the workbook as a new file. Includes robust error handling for missing files and save failures.
class UpdatePowerQueryCsvPath
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains Power Query formulas
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // New CSV file path to be referenced in the Power Query formula
            string newCsvPath = @"C:\Data\newsource.csv";

            // Ensure the workbook actually contains a DataMashup object with Power Query formulas
            if (workbook.DataMashup != null && workbook.DataMashup.PowerQueryFormulas != null)
            {
                // Iterate through all Power Query formulas
                foreach (var formulaObj in workbook.DataMashup.PowerQueryFormulas)
                {
                    try
                    {
                        // Use dynamic to access properties without needing the exact type at compile time
                        dynamic formula = formulaObj;

                        // Retrieve the current M expression
                        string currentFormula = formula.FormulaText as string;

                        if (string.IsNullOrEmpty(currentFormula))
                            continue;

                        // Replace any quoted file path inside the M expression with the new path
                        // Example original value: Csv.Document(File.Contents("C:\\Old\\path.csv"))
                        string updatedFormula = Regex.Replace(
                            currentFormula,
                            "\"[^\"]+\"",
                            $"\"{newCsvPath}\"",
                            RegexOptions.IgnoreCase);

                        // Assign the updated M expression back to the formula
                        formula.FormulaText = updatedFormula;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to update a Power Query formula: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain any Power Query formulas.");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with updated Power Query source: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
