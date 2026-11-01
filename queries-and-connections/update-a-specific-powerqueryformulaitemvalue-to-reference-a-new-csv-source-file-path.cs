// Title: How to update a Power Query formula to reference a new CSV file in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# change power query csv source path using aspose.cells | edit PowerQueryFormulaItem value in an existing workbook with Aspose.Cells | update M code for Power Query in an Excel file programmatically .NET | replace CSV file reference in Power Query formula via Aspose.Cells | modify Power Query source file path across worksheets using C#
// Tags: Aspose.Cells update Power Query CSV source | C# modify PowerQueryFormulaItem value | dynamic access PowerQueryFormulaCollection | replace CSV path in Excel workbook | programmatic M code formula update

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

namespace AsposeCellsPowerQueryUpdate
{
    // The example loads an existing .xlsx workbook, iterates through each worksheet to locate a Power Query named "MyQuery" via the dynamic PowerQueryFormulaCollection, builds a new M‑code formula that points to a specified CSV file, assigns this string to the PowerQueryFormulaItem.Value, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputFile = @"C:\Path\To\Your\Workbook.xlsx";
            string outputFile = @"C:\Path\To\Your\ModifiedWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the existing workbook
                var workbook = new Workbook(inputFile);

                // Define the Power Query name to update and the new CSV source path
                string targetQueryName = "MyQuery"; // replace with the actual query name
                string newCsvPath = @"C:\Data\NewSource.csv";

                // Verify that the new CSV file exists
                if (!File.Exists(newCsvPath))
                {
                    Console.WriteLine($"CSV source file not found: {newCsvPath}");
                    return;
                }

                // Iterate through all worksheets to locate Power Query formulas
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Use dynamic to access PowerQueryFormulaCollection (may not exist in older versions)
                        dynamic dynSheet = sheet;
                        var pqCollection = dynSheet.PowerQueryFormulaCollection as System.Collections.IEnumerable;
                        if (pqCollection == null) continue;

                        foreach (dynamic item in pqCollection)
                        {
                            // Match the query by name (case‑insensitive)
                            string name = item.Name as string;
                            if (string.Equals(name, targetQueryName, StringComparison.OrdinalIgnoreCase))
                            {
                                // Build a new Power Query formula pointing to the new CSV file
                                string newFormula = $@"let
    Source = Csv.Document(File.Contents(""{newCsvPath}""),[Delimiter="","", Columns=5, Encoding=1252, QuoteStyle=QuoteStyle.None])
in
    Source";

                                // Update the formula value
                                item.Value = newFormula;
                                Console.WriteLine($"Updated Power Query '{targetQueryName}' in worksheet '{sheet.Name}'.");
                                break; // Assuming only one match per worksheet
                            }
                        }
                    }
                    catch (RuntimeBinderException)
                    {
                        // Power Query API not available in this version; skip this worksheet
                        continue;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
