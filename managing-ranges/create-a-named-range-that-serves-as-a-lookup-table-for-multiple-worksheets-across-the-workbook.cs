// Title: Define a Global Named Range for Lookup and Use VLOOKUP Across Sheets (Aspose.Cells C#)
// Description: Creates a workbook, builds a two‑column lookup table on the first sheet, defines a global named range "LookupTable" (A1:B4), applies VLOOKUP formulas on a second sheet that reference this range, calculates the formulas, retrieves the range with GetRangeByName to display its address and first value, and saves the file.
// Keywords: Aspose.Cells named range | C# global named range | lookup table Excel | VLOOKUP named range Aspose | GetRangeByName .NET | multi‑sheet lookup Aspose.Cells | create named range C#
// Common Searches: Aspose.Cells create named range for lookup | How to use VLOOKUP with a named range in Aspose.Cells C# | GetRangeByName example Aspose.Cells .NET | Reference a global named range across worksheets Aspose.Cells | Define and retrieve named range in C# workbook
// Developer Intent: Create a named range that holds lookup data and reference it in formulas on other worksheets.
// Use Cases: Expose a two‑column table as a named range for reuse in multiple sheets | Apply VLOOKUP formulas that point to the named range to fetch related values | Programmatically retrieve the named range to verify its address or first cell value | Calculate formulas automatically and save the workbook with lookup results
// AI Prompts: Generate C# code that creates a named range 'LookupTable' covering A1:B4 using Aspose.Cells. | Show how to write VLOOKUP formulas that reference a global named range on another worksheet. | Provide a snippet that calls GetRangeByName to obtain a named range and prints its address and first cell value. | Explain how to calculate formulas and save the workbook after using a named range for lookup.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeLookupExample
{
    // Creates a workbook, builds a two‑column lookup table on the first sheet, defines a global named range "LookupTable" (A1:B4), applies VLOOKUP formulas on a second sheet that reference this range, calculates the formulas, retrieves the range with GetRangeByName to display its address and first value, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare lookup data on the first worksheet
                // -------------------------------------------------
                Worksheet lookupSheet = workbook.Worksheets[0];
                Cells lookupCells = lookupSheet.Cells;

                // Header
                lookupCells["A1"].PutValue("Item");
                lookupCells["B1"].PutValue("Price");

                // Sample data
                lookupCells["A2"].PutValue("Apple");
                lookupCells["B2"].PutValue(1.20);
                lookupCells["A3"].PutValue("Banana");
                lookupCells["B3"].PutValue(0.80);
                lookupCells["A4"].PutValue("Cherry");
                lookupCells["B4"].PutValue(2.50);

                // -------------------------------------------------
                // 2. Create a named range that covers the lookup table
                // -------------------------------------------------
                AsposeRange namedRange = lookupCells.CreateRange("A1", "B4");
                namedRange.Name = "LookupTable";

                // -------------------------------------------------
                // 3. Use the named range in another worksheet (lookup)
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets.Add("Data");
                Cells dataCells = dataSheet.Cells;

                // Input items to lookup
                dataCells["A1"].PutValue("Item");
                dataCells["B1"].PutValue("Price (Lookup)");
                dataCells["A2"].PutValue("Apple");
                dataCells["A3"].PutValue("Cherry");
                dataCells["A4"].PutValue("Banana");

                // Apply VLOOKUP formula that references the named range
                dataCells["B2"].Formula = "=VLOOKUP(A2,LookupTable,2,FALSE)";
                dataCells["B3"].Formula = "=VLOOKUP(A3,LookupTable,2,FALSE)";
                dataCells["B4"].Formula = "=VLOOKUP(A4,LookupTable,2,FALSE)";

                // Calculate formulas so that values are materialized
                workbook.CalculateFormula();

                // -------------------------------------------------
                // 4. Retrieve the named range using GetRangeByName (global scope)
                // -------------------------------------------------
                AsposeRange retrievedRange = workbook.Worksheets.GetRangeByName("LookupTable");
                if (retrievedRange != null)
                {
                    Console.WriteLine("Named range found: " + retrievedRange.Address);
                    Console.WriteLine("First cell value: " + retrievedRange[0, 0].StringValue);
                }
                else
                {
                    Console.WriteLine("Named range not found.");
                }

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("NamedRangeLookupDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
