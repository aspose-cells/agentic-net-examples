// Title: C# Unit Test: StyleFlag.QuotePrefix = false Preserves Existing Quote‑Prefix in Aspose.Cells
// Description: Demonstrates how to verify that setting StyleFlag.QuotePrefix to false does not modify a cell's existing QuotePrefix flag. The test creates a workbook, inserts a value with a leading single quote, confirms the flag is true, applies a new style with the flag disabled, saves to a memory stream, reloads the workbook, and asserts the flag remains true.
// Keywords: Aspose.Cells | StyleFlag | QuotePrefix | unit test | .NET | C# | leading single quote | cell style flag | save load persistence | Aspose.Cells StyleFlag false
// Common Searches: Aspose.Cells StyleFlag QuotePrefix false unit test | verify QuotePrefix flag remains after applying style in C# | preserve leading single quote Aspose.Cells after style change | C# test QuotePrefix persistence after workbook save | how to assert QuotePrefix flag in Aspose.Cells unit test
// Developer Intent: Confirm that StyleFlag.QuotePrefix set to false leaves an existing QuotePrefix flag unchanged.
// Use Cases: Automated regression test to ensure applying a style without QuotePrefix does not clear existing leading‑quote formatting. | Validate that the QuotePrefix flag survives workbook serialization and deserialization. | Guarantee consistent behavior when updating cell styles in bulk operations.
// AI Prompts: Generate an MSTest method that asserts StyleFlag.QuotePrefix = false does not affect a cell's QuotePrefix flag in Aspose.Cells for .NET. | Write an xUnit test verifying QuotePrefix persistence after saving and loading a workbook using Aspose.Cells C# API. | Provide a NUnit example that checks the QuotePrefix flag remains true when a style is applied with StyleFlag.QuotePrefix set to false.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates how to verify that setting StyleFlag.QuotePrefix to false does not modify a cell's existing QuotePrefix flag. The test creates a workbook, inserts a value with a leading single quote, confirms the flag is true, applies a new style with the flag disabled, saves to a memory stream, reloads the workbook, and asserts the flag remains true.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put a value that starts with a single quote.
                // Aspose.Cells treats the leading quote as a formatting flag.
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("'12345");

                // Verify that the cell's style has QuotePrefix set to true initially.
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("Initial QuotePrefix should be true.");

                // Create a style with QuotePrefix set to false (default) and a StyleFlag with QuotePrefix false.
                Style style = workbook.CreateStyle();
                style.QuotePrefix = false; // Explicitly set for clarity.

                StyleFlag flag = new StyleFlag();
                flag.QuotePrefix = false; // Ensure the flag does not apply QuotePrefix changes.

                // Apply the style using the flag. Since the flag is false, the existing QuotePrefix should remain unchanged.
                cell.SetStyle(style, flag);

                // Assert that the QuotePrefix is still true after applying the style with the flag set to false.
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should remain true when StyleFlag.QuotePrefix is false.");

                // Save the workbook to a memory stream to test persistence.
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0;

                    // Load the workbook from the memory stream.
                    Workbook loadedWorkbook = new Workbook(ms);
                    Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];

                    // Verify that the loaded cell still retains the QuotePrefix flag.
                    if (!loadedCell.GetStyle().QuotePrefix)
                        throw new Exception("Loaded cell should retain QuotePrefix after save/load.");
                }

                Console.WriteLine("All checks passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
