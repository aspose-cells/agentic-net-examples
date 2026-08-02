// Title: C# Unit Test for Aspose.Cells FreezePanes – Verify Frozen Row Count
// Description: Demonstrates how to create an automated test (NUnit or MSTest) that applies Worksheet.FreezePanes to a workbook, then uses GetFreezedPanes to assert that the frozen rows, columns and start positions match the expected values.
// Keywords: Aspose.Cells | FreezePanes | C# unit test | NUnit | MSTest | GetFreezedPanes | worksheet freeze panes | Excel API testing | automated verification | Aspose.Cells example
// Common Searches: Aspose.Cells FreezePanes unit test C# | How to assert frozen rows with GetFreezedPanes | NUnit test for worksheet freeze panes | Validate FreezePanes parameters in Aspose.Cells | MSTest example for FreezePanes verification
// Developer Intent: Create an automated test that confirms FreezePanes correctly freezes the specified number of rows (and columns) in a worksheet.
// Use Cases: Validate that FreezePanes(5, 0, 5, 0) sets the correct start row, column, and frozen row count. | Integrate the test into a CI pipeline to detect regressions in pane‑freezing behavior. | Use the pattern as a template for testing other FreezePanes scenarios, such as column freezing or combined row/column freezes.
// AI Prompts: Generate an NUnit test method that creates a Workbook, calls FreezePanes(5,0,5,0) on the first worksheet, and uses Assert.AreEqual to verify the values returned by GetFreezedPanes. | Provide a MSTest example that checks the FreezePanes parameters and fails with a clear message if any value is incorrect. | Write a reusable helper function for Aspose.Cells tests that validates frozen pane settings given expected row, column, rows, and columns.

using System;
using Aspose.Cells;

// Demonstrates how to create an automated test (NUnit or MSTest) that applies Worksheet.FreezePanes to a workbook, then uses GetFreezedPanes to assert that the frozen rows, columns and start positions match the expected values.
public class FreezePanesDemo
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze panes: start at row index 5 (6th row), column index 0,
            // with 5 visible frozen rows and 0 frozen columns
            int freezeRow = 5;
            int freezeColumn = 0;
            int frozenRows = 5;
            int frozenColumns = 0;
            worksheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);

            // Verify that the worksheet reports frozen panes and the parameters match
            bool hasFreeze = worksheet.GetFreezedPanes(out int row, out int column, out int rows, out int columns);
            if (!hasFreeze)
                throw new InvalidOperationException("Worksheet should indicate that panes are frozen.");

            if (row != freezeRow)
                throw new InvalidOperationException($"Freeze position row does not match. Expected {freezeRow}, got {row}.");

            if (column != freezeColumn)
                throw new InvalidOperationException($"Freeze position column does not match. Expected {freezeColumn}, got {column}.");

            if (rows != frozenRows)
                throw new InvalidOperationException($"Number of frozen rows does not match. Expected {frozenRows}, got {rows}.");

            if (columns != frozenColumns)
                throw new InvalidOperationException($"Number of frozen columns does not match. Expected {frozenColumns}, got {columns}.");

            Console.WriteLine("Freeze panes applied and verified successfully.");

            // Optional: save the workbook if you want to inspect the file manually
            // workbook.Save("FreezePanesDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
