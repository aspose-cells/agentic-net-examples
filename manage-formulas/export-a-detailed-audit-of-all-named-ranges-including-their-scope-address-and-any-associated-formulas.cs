// Title: Generate a CSV report of all named ranges in an Excel workbook, including scope, address, and formula, using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, iterates through workbook.Worksheets.Names, extracts each name's scope, address, and RefersTo formula, and writes the collected data to a CSV file. | Enhance the example to also read any comment attached to a named range and add a Comment column to the CSV export. | Create a reusable method that returns a collection of objects containing Name, Scope, Address, and Formula for every defined name in a workbook loaded with Aspose.Cells.
// Common Searches: how to list all named ranges with scope and formula using Aspose.Cells C# | export named range details to CSV in Aspose.Cells .NET | retrieve address of a defined name programmatically with Aspose.Cells | Aspose.Cells C# get worksheet‑specific named range scope | C# generate audit of Excel named ranges using Aspose.Cells
// Tags: Aspose.Cells export named ranges CSV | Aspose.Cells get named range scope | Aspose.Cells retrieve defined name address | Aspose.Cells list workbook names | named range audit .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeAudit
{
    // The sample checks for the input Excel file, loads it with Aspose.Cells, and loops through all defined names in workbook.Worksheets.Names. For each name it obtains the identifier via reflection, determines whether the scope is workbook‑wide or tied to a specific worksheet, attempts to get the associated range to capture its address, and records the RefersTo formula. All information (Name, Scope, Address, Formula) is CSV‑escaped and written to NamedRangeAudit.csv.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "NamedRangeAudit.csv";

            // Verify input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            Workbook workbook;
            try
            {
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            try
            {
                using (var writer = new StreamWriter(outputFile))
                {
                    // CSV header
                    writer.WriteLine("Name,Scope,Address,Formula");

                    // Iterate through all defined names
                    foreach (var item in workbook.Worksheets.Names)
                    {
                        // Cast to Aspose.Cells.Name (may be null if type differs)
                        var definedName = item as Aspose.Cells.Name;
                        if (definedName == null)
                            continue;

                        // Retrieve the name string via reflection to avoid version issues
                        string name = string.Empty;
                        var nameProp = definedName.GetType().GetProperty("Name");
                        if (nameProp != null)
                        {
                            var val = nameProp.GetValue(definedName);
                            name = val?.ToString() ?? string.Empty;
                        }

                        // Determine scope: workbook or specific worksheet
                        string scope = "Workbook";
                        var wsProp = definedName.GetType().GetProperty("Worksheet");
                        if (wsProp != null)
                        {
                            var ws = wsProp.GetValue(definedName) as Worksheet;
                            if (ws != null)
                            {
                                scope = ws.Name;
                            }
                        }

                        // Try to obtain the range the name refers to
                        AsposeRange rangeObj = null;
                        try
                        {
                            rangeObj = definedName.GetRange();
                        }
                        catch
                        {
                            // Ignored – the name may refer to a constant or formula
                        }

                        // Address of the range, if available
                        string address = string.Empty;
                        if (rangeObj != null)
                        {
                            try
                            {
                                address = rangeObj.Address;
                            }
                            catch
                            {
                                // Ignored – fallback to empty address
                            }
                        }

                        // Formula or reference (including leading '=')
                        string formula = definedName.RefersTo ?? string.Empty;

                        // Escape fields for CSV
                        string csvName = $"\"{name}\"";
                        string csvScope = $"\"{scope}\"";
                        string csvAddress = $"\"{address}\"";
                        string csvFormula = $"\"{formula}\"";

                        writer.WriteLine($"{csvName},{csvScope},{csvAddress},{csvFormula}");
                    }
                }

                Console.WriteLine($"Named range audit has been exported to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }
}
