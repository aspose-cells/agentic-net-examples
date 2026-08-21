// Title: Aspose.Cells .NET: Preserve QuotePrefix after applying NumberFormat with StyleFlag
// Description: C# example that creates a workbook, sets cell A1 to a text value with QuotePrefix enabled, updates only the number format using StyleFlag, and verifies that the QuotePrefix flag remains true before and after the style change, after saving to XLSX, and after reloading the file.
// Keywords: Aspose.Cells | QuotePrefix | StyleFlag | NumberFormat | C# | .NET | preserve leading apostrophe | SetStyle | Excel cell formatting | save and reload workbook | XLSX
// Common Searches: Aspose.Cells keep QuotePrefix after style change | SetStyle with StyleFlag preserve leading quote | QuotePrefix true after saving workbook Aspose.Cells | C# Aspose.Cells update number format without losing QuotePrefix | How to use StyleFlag to change only NumberFormat in Aspose.Cells
// Developer Intent: Confirm that a cell’s QuotePrefix stays true when its number format is modified with a StyleFlag and after the workbook is saved and reloaded.
// Use Cases: Validate that applying a NumberFormat style via StyleFlag does not reset QuotePrefix. | Ensure leading apostrophe remains after serializing an Aspose.Cells workbook to XLSX. | Demonstrate selective style updates (NumberFormat only) while preserving other cell style attributes.
// AI Prompts: Generate a C# unit test using Aspose.Cells that asserts QuotePrefix is unchanged after applying a NumberFormat StyleFlag. | Provide sample code to modify only the number format of a cell without affecting its QuotePrefix, then verify the property after saving and loading the file. | Explain how StyleFlag works with SetStyle to keep unchanged style properties such as QuotePrefix in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixValidation
{
    // C# example that creates a workbook, sets cell A1 to a text value with QuotePrefix enabled, updates only the number format using StyleFlag, and verifies that the QuotePrefix flag remains true before and after the style change, after saving to XLSX, and after reloading the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Step 1: Set a cell value that looks like a number but should be
            //         treated as text with a leading quote (QuotePrefix = true)
            // ------------------------------------------------------------
            Cell cell = cells["A1"];
            cell.PutValue("123456");               // Put a numeric string
            Style initialStyle = workbook.CreateStyle();
            initialStyle.QuotePrefix = true;       // Enable QuotePrefix
            cell.SetStyle(initialStyle);           // Apply the style

            // Verify initial QuotePrefix
            Console.WriteLine("Initial QuotePrefix: " + cell.GetStyle().QuotePrefix); // Expected: True

            // ------------------------------------------------------------
            // Step 2: Create a style that changes only the number format.
            //         Use a StyleFlag to apply only the NumberFormat property.
            // ------------------------------------------------------------
            Style numberFormatStyle = workbook.CreateStyle();
            numberFormatStyle.Custom = "#,##0.00"; // Example custom number format

            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;              // Apply only number format

            // Apply the style with the flag. QuotePrefix flag is NOT set,
            // so QuotePrefix should remain unchanged.
            cell.SetStyle(numberFormatStyle, flag);

            // ------------------------------------------------------------
            // Step 3: Validate that QuotePrefix is still true after the update.
            // ------------------------------------------------------------
            bool quotePrefixAfterUpdate = cell.GetStyle().QuotePrefix;
            Console.WriteLine("QuotePrefix after NumberFormat update: " + quotePrefixAfterUpdate); // Expected: True

            // ------------------------------------------------------------
            // Step 4: Save the workbook, reload it, and verify QuotePrefix again.
            // ------------------------------------------------------------
            string filePath = "QuotePrefixValidation.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the saved workbook
            Workbook loadedWorkbook = new Workbook(filePath);
            Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];
            bool quotePrefixAfterLoad = loadedCell.GetStyle().QuotePrefix;
            Console.WriteLine("QuotePrefix after reload: " + quotePrefixAfterLoad); // Expected: True
        }
    }
}
