// Title: Check VLOOKUP support in Aspose.Cells for .NET with HasCustomFunction
// Description: This C# example creates a workbook, adds a simple table, inserts a VLOOKUP formula, calculates the sheet, and uses the HasCustomFunction property to confirm that VLOOKUP is recognized as a native function (not custom) according to Aspose.Cells supported‑functions documentation.
// Keywords: Aspose.Cells VLOOKUP .NET | HasCustomFunction C# | supported Excel functions Aspose | detect unsupported formulas | Aspose.Cells formula validation
// Common Searches: Is VLOOKUP supported by Aspose.Cells | Aspose.Cells HasCustomFunction usage | check unsupported Excel functions in .NET | Aspose.Cells supported functions list | how to verify formula compatibility Aspose
// Developer Intent: Confirm whether the VLOOKUP function is natively supported by Aspose.Cells for .NET.
// Use Cases: Validate workbook formulas before publishing to ensure no unsupported functions are present. | Automate regression tests that flag cells marked as custom functions after calculation. | Generate dynamic reports that rely on VLOOKUP without risking runtime errors.
// AI Prompts: Create a C# routine that scans all cells in an Aspose.Cells workbook and returns those where HasCustomFunction is true. | Write a unit test in .NET that asserts a VLOOKUP formula is not flagged as a custom function after wb.CalculateFormula(). | Provide code that reads the Aspose.Cells supported‑functions documentation and programmatically verifies if a given function name (e.g., VLOOKUP) is supported.

using System;
using Aspose.Cells;

namespace AsposeCellsVLookupCheck
{
    // This C# example creates a workbook, adds a simple table, inserts a VLOOKUP formula, calculates the sheet, and uses the HasCustomFunction property to confirm that VLOOKUP is recognized as a native function (not custom) according to Aspose.Cells supported‑functions documentation.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a simple table for VLOOKUP
            // Header
            cells["A1"].PutValue("Key");
            cells["B1"].PutValue("Value");
            // Data rows
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue(30);

            // Set a VLOOKUP formula that should be supported
            // Lookup "Banana" in the table A1:B4 and return the value from the second column
            cells["D1"].Formula = "=VLOOKUP(\"Banana\", A1:B4, 2, FALSE)";

            // Calculate formulas
            wb.CalculateFormula();

            // Check if the formula used a custom (unsupported) function
            bool hasCustom = cells["D1"].HasCustomFunction;

            // Output the result
            Console.WriteLine("VLOOKUP formula result: " + cells["D1"].StringValue);
            Console.WriteLine("Is VLOOKUP a custom/unsupported function? " + hasCustom);
            Console.WriteLine("Conclusion: VLOOKUP is " + (hasCustom ? "NOT supported" : "supported") + " by Aspose.Cells.");

            // Optionally save the workbook for visual verification
            wb.Save("VLookupCheck.xlsx");
        }
    }
}
