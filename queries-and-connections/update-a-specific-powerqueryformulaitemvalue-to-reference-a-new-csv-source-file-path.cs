// Title: How to programmatically change the Power Query CSV source path in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file, finds the PowerQueryFormulaItem named "Source", and replaces its file path with a new CSV location. | Show how to iterate through workbook.DataMashup.PowerQueryFormulas, update the Source expression to point to a different CSV file, and save the modified workbook. | Create a reusable method UpdatePowerQueryCsvSource(string workbookPath, string oldCsvPath, string newCsvPath) that performs the source‑path replacement and returns the path of the saved file.
// Common Searches: aspnet c# update power query source file path in existing excel workbook using aspose.cells | replace old csv path with new csv in PowerQueryFormulaItem via Aspose.Cells .NET | programmatically modify Power Query connection string in an .xlsx file using C# | how to change the Source item of a Power Query formula in Aspose.Cells | update Power Query CSV source without opening Excel UI using Aspose.Cells
// Tags: Aspose.Cells PowerQueryFormulaItem update | C# change Power Query CSV path | Excel workbook Power Query source modification | Aspose.Cells .NET Power Query editing | programmatic Power Query connection update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryUpdate
{
    // Loads an existing Excel workbook, locates the Power Query formula item named "Source", replaces the old CSV file path with a new one, and saves the workbook with the updated Power Query source.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the existing workbook that contains Power Query formulas
                string sourceWorkbookPath = "input.xlsx";

                // Verify the source workbook exists
                if (!File.Exists(sourceWorkbookPath))
                {
                    Console.WriteLine($"Source workbook not found: {sourceWorkbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceWorkbookPath);

                // Define the new CSV file path you want the Power Query to reference
                string newCsvPath = @"C:\Data\NewSourceFile.csv";

                // Flag to indicate whether the target item was found
                bool itemUpdated = false;

                // Ensure the workbook actually contains Power Query data
                if (workbook.DataMashup != null && workbook.DataMashup.PowerQueryFormulas != null)
                {
                    // Iterate through all Power Query formulas in the workbook
                    foreach (PowerQueryFormula formula in workbook.DataMashup.PowerQueryFormulas)
                    {
                        // Iterate through the items of each formula
                        foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                        {
                            // Typically the item that holds the source path is named "Source"
                            if (item.Name.Equals("Source", StringComparison.OrdinalIgnoreCase))
                            {
                                // Replace any existing file path with the new CSV path
                                if (!string.IsNullOrEmpty(item.Value))
                                {
                                    item.Value = item.Value.Replace(@"C:\OldPath\OldFile.csv", newCsvPath);
                                }

                                // Optionally set the whole expression directly (uncomment if needed)
                                // item.Value = $"Source=Csv.Document(File.Contents(\"{newCsvPath}\"),[Delimiter=\",\", Columns=5, Encoding=1252, QuoteStyle=QuoteStyle.None])";

                                itemUpdated = true;
                                break; // Assuming only one "Source" item per formula
                            }
                        }

                        if (itemUpdated) break;
                    }
                }
                else
                {
                    Console.WriteLine("The workbook does not contain any Power Query formulas.");
                }

                if (!itemUpdated)
                {
                    Console.WriteLine("No PowerQueryFormulaItem named 'Source' was found in the workbook.");
                }
                else
                {
                    // Save the modified workbook
                    string outputWorkbookPath = "output.xlsx";
                    workbook.Save(outputWorkbookPath);
                    Console.WriteLine($"Workbook saved with updated Power Query source: {outputWorkbookPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
