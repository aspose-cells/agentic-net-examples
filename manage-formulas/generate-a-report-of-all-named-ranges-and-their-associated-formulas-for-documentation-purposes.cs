// Title: C# – Generate a Named‑Range and Formula Report with Aspose.Cells
// Description: A console program that loads an Excel workbook using Aspose.Cells for .NET, enumerates all defined names, extracts each name’s RefersTo formula, resolves the actual cell address when possible, and prints the results. The workbook can be saved after inspection.
// Keywords: Aspose.Cells | C# named ranges | list defined names | Excel named range report | GetRange Aspose.Cells | RefersTo formula | .NET Excel automation | named range documentation | workbook NameCollection
// Common Searches: list all named ranges with Aspose.Cells C# | how to get RefersTo formula of a defined name in .NET | retrieve range address from a named range using Aspose.Cells | C# code to export named range definitions from Excel | Aspose.Cells example for named‑range documentation
// Developer Intent: Build a .NET console app that reads an Excel file and outputs every named range together with its formula and resolved address.
// Use Cases: Create documentation of all named ranges for audit or compliance. | Validate that each defined name points to a valid cell range before distribution. | Export named‑range metadata to CSV/JSON for integration with reporting tools.
// AI Prompts: Write C# code with Aspose.Cells that reads a workbook and writes a CSV containing Name, RefersTo formula, and resolved address for each defined name. | Provide a method returning Dictionary<string, string> where the key is the named range and the value is its address, handling GetRange exceptions gracefully. | Generate a PowerShell script that calls a compiled .NET assembly to produce a named‑range report for a given Excel file.

using System;
using System.IO;
using Aspose.Cells;

namespace NamedRangesReport
{
    // A console program that loads an Excel workbook using Aspose.Cells for .NET, enumerates all defined names, extracts each name’s RefersTo formula, resolves the actual cell address when possible, and prints the results. The workbook can be saved after inspection.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains the named ranges.
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook.
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // Access the collection of all defined names in the workbook.
            NameCollection names = workbook.Worksheets.Names;

            Console.WriteLine("Named Ranges Report");
            Console.WriteLine("====================");
            Console.WriteLine();

            // Iterate through each Name object.
            foreach (Name name in names)
            {
                // Name text (e.g., "MyRange")
                string nameText = name.Text;

                // The formula that the name refers to (starts with '=').
                string refersToFormula = name.RefersTo;

                // Try to obtain the actual Range object if the name refers to a range.
                // GetRange may return null for external references or non‑range definitions.
                string rangeAddress = "N/A";
                try
                {
                    Aspose.Cells.Range range = name.GetRange();
                    if (range != null)
                    {
                        rangeAddress = range.Address;
                    }
                }
                catch
                {
                    // Ignored – GetRange can throw if the reference is not a simple range.
                }

                // Output the information.
                Console.WriteLine($"Name          : {nameText}");
                Console.WriteLine($"Refers To     : {refersToFormula}");
                Console.WriteLine($"Range Address : {rangeAddress}");
                Console.WriteLine(new string('-', 40));
            }

            // Optionally save the workbook if any changes were made.
            try
            {
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved as \"output.xlsx\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
