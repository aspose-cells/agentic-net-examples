// Title: Check if Aspose.Cells for .NET Supports the VLOOKUP Function
// Description: A C# sample that creates a workbook, builds a small lookup table, inserts a VLOOKUP formula, runs calculation, and inspects the result cell for errors or custom‑function flags to confirm whether the VLOOKUP function is available in the current Aspose.Cells version.
// Keywords: Aspose.Cells VLOOKUP support | C# VLOOKUP formula evaluation | Aspose.Cells supported Excel functions | detect unsupported Excel functions | calculate formula Aspose.Cells .NET | VLOOKUP error handling Aspose
// Common Searches: does Aspose.Cells support VLOOKUP | how to test VLOOKUP in Aspose.Cells .NET | Aspose.Cells VLOOKUP returns error | verify Excel function support Aspose.Cells | check if VLOOKUP works in Aspose.Cells
// Developer Intent: Confirm whether the VLOOKUP function can be calculated by Aspose.Cells for .NET and obtain a boolean status at runtime.
// Use Cases: Validate compatibility of incoming workbooks that rely on VLOOKUP before bulk processing. | Implement a fallback routine when VLOOKUP is unavailable, preventing calculation failures. | Log function‑support metrics for compliance reports or version‑upgrade decisions.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect VLOOKUP support and switches to a manual lookup if needed. | Write a unit test that asserts the VLOOKUP formula evaluates without error in Aspose.Cells. | Explain how Cell.HasCustomFunction and CellValueType.IsError can be used to identify unsupported Excel functions in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsVlookupSupportCheck
{
    // A C# sample that creates a workbook, builds a small lookup table, inserts a VLOOKUP formula, runs calculation, and inspects the result cell for errors or custom‑function flags to confirm whether the VLOOKUP function is available in the current Aspose.Cells version.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook wb = new Workbook();

                // Access the first worksheet
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a simple lookup table (A1:B4)
                cells["A1"].PutValue("Key");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(30);

                // Set a VLOOKUP formula that searches for "Banana"
                // Syntax: VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
                cells["D1"].Formula = "=VLOOKUP(\"Banana\", A1:B4, 2, FALSE)";

                // Calculate formulas (this will attempt to evaluate VLOOKUP)
                wb.CalculateFormula();

                // Retrieve the result
                Cell resultCell = cells["D1"];
                string resultText = resultCell.StringValue;

                // Determine support:
                // If VLOOKUP is unsupported, Aspose.Cells marks the cell as having a custom function
                // or returns an error value. We check both conditions.
                bool isSupported = !resultCell.HasCustomFunction &&
                                   resultCell.Type != CellValueType.IsError;

                Console.WriteLine($"VLOOKUP supported: {isSupported}");
                Console.WriteLine($"Result of VLOOKUP: {resultText}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
