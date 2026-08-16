// Title: Bulk replace VLOOKUP with XLOOKUP in Excel workbooks using Aspose.Cells for .NET
// Description: Loads an Excel file (or creates a sample workbook), scans every worksheet for cells that contain the VLOOKUP function, swaps the function name to XLOOKUP while keeping the original arguments, recalculates all formulas, and saves the updated workbook. Demonstrates FindOptions for formula‑only search and shows error handling in C#.
// Keywords: Aspose.Cells | C# | .NET | XLOOKUP | VLOOKUP | Excel formula conversion | bulk formula update | FindOptions OnlyFormulas | Excel automation | legacy lookup replacement | Excel 365 compatibility | programmatic workbook editing
// Common Searches: replace VLOOKUP with XLOOKUP using Aspose.Cells | search and modify Excel formulas in .NET | bulk update lookup functions in multiple worksheets | convert legacy VLOOKUP to XLOOKUP programmatically | Aspose.Cells find formulas only option
// Developer Intent: Programmatically change every VLOOKUP formula in a workbook to an XLOOKUP formula with Aspose.Cells.
// Use Cases: Modernize legacy spreadsheets before sharing with users of newer Excel versions. | Automate bulk migration of lookup functions across large document libraries. | Ensure accurate calculations after conversion by triggering a full workbook recalculation.
// AI Prompts: Generate C# code that locates all VLOOKUP formulas in an Excel workbook and replaces them with XLOOKUP using Aspose.Cells. | Create a robust routine that parses VLOOKUP arguments and builds equivalent XLOOKUP syntax, handling optional parameters and errors. | Show an example that builds a sample workbook, inserts a VLOOKUP formula, converts it to XLOOKUP, recalculates, and saves the file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file (or creates a sample workbook), scans every worksheet for cells that contain the VLOOKUP function, swaps the function name to XLOOKUP while keeping the original arguments, recalculates all formulas, and saves the updated workbook. Demonstrates FindOptions for formula‑only search and shows error handling in C#.
    public class ReplaceVlookupWithXlookup
    {
        public static void Run()
        {
            try
            {
                // Input workbook path
                string inputPath = "LegacyVlookupWorkbook.xlsx";

                // Ensure the input file exists; create a sample workbook if missing
                if (!File.Exists(inputPath))
                {
                    Workbook sampleWb = new Workbook();
                    Worksheet ws = sampleWb.Worksheets[0];
                    ws.Name = "SampleSheet";

                    // Add sample data
                    ws.Cells["A1"].PutValue("Key");
                    ws.Cells["B1"].PutValue("Value");
                    ws.Cells["A2"].PutValue("Item1");
                    ws.Cells["B2"].PutValue(100);
                    ws.Cells["A3"].PutValue("Item2");
                    ws.Cells["B3"].PutValue(200);

                    // Add a VLOOKUP formula that will be replaced
                    ws.Cells["C2"].Formula = "=VLOOKUP(A2,A1:B3,2,FALSE)";

                    sampleWb.Save(inputPath);
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Define find options to search only in formulas and allow partial matches
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.OnlyFormulas,
                    LookAtType = LookAtType.Contains
                };

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Find the first occurrence of a VLOOKUP formula
                    Cell foundCell = sheet.Cells.Find("VLOOKUP", null, findOptions);

                    // Continue searching until no more VLOOKUP formulas are found
                    while (foundCell != null)
                    {
                        // Get the original formula
                        string originalFormula = foundCell.Formula;

                        // Simple conversion: replace the function name while keeping the argument list
                        string updatedFormula = originalFormula.Replace("VLOOKUP(", "XLOOKUP(");

                        // Apply the new formula to the cell
                        foundCell.Formula = updatedFormula;

                        // Search for the next VLOOKUP formula starting after the current cell
                        foundCell = sheet.Cells.Find("VLOOKUP", foundCell, findOptions);
                    }
                }

                // Recalculate all formulas to ensure the new XLOOKUP functions are evaluated
                workbook.CalculateFormula();

                // Save the modified workbook
                string outputPath = "WorkbookWithXlookup.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceVlookupWithXlookup.Run();
        }
    }
}
