// Title: Define a Workbook‑Scoped Named Range for VLOOKUP Across Sheets with Aspose.Cells (C#)
// Description: Creates a workbook, adds a "Lookup" sheet with key/value pairs, defines a global named range "LookupTable", inserts VLOOKUP formulas on a "Data" sheet that reference the named range, calculates formulas, retrieves the range via GetRangeByName, and saves the file.
// Keywords: Aspose.Cells named range | global named range C# | VLOOKUP across worksheets | GetRangeByName example | Aspose.Cells .NET lookup table | Create workbook‑level name | Excel formula calculation Aspose
// Common Searches: Aspose.Cells define workbook level named range | use VLOOKUP with named range in Aspose.Cells | retrieve named range by name Aspose.Cells C# | calculate formulas after setting VLOOKUP Aspose | save workbook after formula evaluation Aspose.Cells
// Developer Intent: Create a workbook‑scoped named range for a lookup table and reference it in VLOOKUP formulas on other worksheets.
// Use Cases: Build a dedicated lookup sheet, populate keys and values, and assign a global name for reuse. | Apply VLOOKUP formulas on separate sheets that point to the named range, enabling centralized data maintenance. | Programmatically fetch the named range with GetRangeByName to verify or modify its address. | Trigger full formula calculation and persist the resolved results in the saved workbook.
// AI Prompts: Generate C# code that defines a workbook‑level named range for a two‑column lookup table and uses it in VLOOKUP formulas on another sheet with Aspose.Cells. | Show how to retrieve and update a named range by name in an existing Aspose.Cells workbook. | Explain how to force calculation of all formulas after inserting VLOOKUP references that use a named range.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeLookup
{
    // Creates a workbook, adds a "Lookup" sheet with key/value pairs, defines a global named range "LookupTable", inserts VLOOKUP formulas on a "Data" sheet that reference the named range, calculates formulas, retrieves the range via GetRangeByName, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -----------------------------------------------------------------
                // 1. Create a worksheet that will hold the lookup table
                // -----------------------------------------------------------------
                Worksheet lookupSheet = workbook.Worksheets[0];
                lookupSheet.Name = "Lookup";

                // Fill the lookup table (A column = keys, B column = values)
                string[] keys = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
                int[] values = { 10, 20, 30, 40, 50 };

                for (int i = 0; i < keys.Length; i++)
                {
                    lookupSheet.Cells[i, 0].PutValue(keys[i]);   // Column A
                    lookupSheet.Cells[i, 1].PutValue(values[i]); // Column B
                }

                // Create a Range object for the lookup table and assign a global name
                AsposeRange lookupRange = lookupSheet.Cells.CreateRange("A1", "B5");
                lookupRange.Name = "LookupTable"; // Global name (Workbook scope)

                // -----------------------------------------------------------------
                // 2. Create another worksheet that will use the lookup table
                // -----------------------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets.Add("Data");

                // Header
                dataSheet.Cells["A1"].PutValue("Fruit");
                dataSheet.Cells["B1"].PutValue("Quantity");

                // Sample fruits to look up
                string[] fruitsToLookup = { "Cherry", "Apple", "Date" };
                for (int i = 0; i < fruitsToLookup.Length; i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(fruitsToLookup[i]); // Column A

                    // Use VLOOKUP with the named range "LookupTable"
                    // Formula: =VLOOKUP(A2,LookupTable,2,FALSE)
                    string formula = $"=VLOOKUP(A{i + 2},LookupTable,2,FALSE)";
                    dataSheet.Cells[i + 1, 1].Formula = formula;
                }

                // Calculate all formulas so that lookup results are materialized
                workbook.CalculateFormula();

                // -----------------------------------------------------------------
                // 3. Demonstrate retrieving the named range via GetRangeByName
                // -----------------------------------------------------------------
                AsposeRange retrievedRange = workbook.Worksheets.GetRangeByName("LookupTable");
                if (retrievedRange != null)
                {
                    Console.WriteLine($"Named range '{retrievedRange.Name}' address: {retrievedRange.RefersTo}");
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "NamedRangeLookupDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
