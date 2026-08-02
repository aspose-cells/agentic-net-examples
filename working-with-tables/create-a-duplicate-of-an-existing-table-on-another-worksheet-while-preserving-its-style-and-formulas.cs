// Title: Copy an Aspose.Cells ListObject to another worksheet while keeping style and formulas (C#)
// Description: Demonstrates how to duplicate a ListObject (Excel table) from one worksheet to another using Aspose.Cells for .NET. The example creates a source table, defines a destination range, applies PasteOptions with KeepOldTables=true, and saves the workbook, ensuring all formatting and formulas are retained.
// Keywords: Aspose.Cells copy table | duplicate ListObject C# | preserve table style Aspose.Cells | keep formulas when copying Excel table | PasteOptions KeepOldTables example | Aspose.Cells range copy across sheets | C# Excel table duplication | Aspose.Cells ListObject clone
// Common Searches: how to copy a ListObject to another sheet in Aspose.Cells | Aspose.Cells preserve table formatting when copying | duplicate Excel table with formulas using Aspose.Cells .NET | PasteOptions KeepOldTables usage | copy Aspose.Cells table to a different worksheet
// Developer Intent: Create an exact copy of an existing ListObject on a new worksheet, retaining its visual style and any embedded formulas.
// Use Cases: Copy a product catalog table to a summary sheet without losing conditional formatting or calculated columns. | Replicate a financial data table across regional reports while preserving complex formulas. | Generate a template where a master table is cloned to each department sheet with identical styling and logic.
// AI Prompts: Show C# code to duplicate an Aspose.Cells ListObject to another worksheet while keeping its style and formulas. | Explain how PasteOptions.KeepOldTables works when copying a table range in Aspose.Cells. | Provide step‑by‑step guidance for cloning an Excel table across sheets using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to duplicate a ListObject (Excel table) from one worksheet to another using Aspose.Cells for .NET. The example creates a source table, defines a destination range, applies PasteOptions with KeepOldTables=true, and saves the workbook, ensuring all formatting and formulas are retained.
class DuplicateTableDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet (source)
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate sample data for the table
            sourceSheet.Cells["A1"].PutValue("Product");
            sourceSheet.Cells["B1"].PutValue("Quantity");
            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("Banana");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("Orange");
            sourceSheet.Cells["B4"].PutValue(15);

            // Create a table (ListObject) that includes the data range (A1:B4)
            int tableIdx = sourceSheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject sourceTable = sourceSheet.ListObjects[tableIdx];
            sourceTable.DisplayName = "ProductsTable";
            sourceTable.TableStyleType = TableStyleType.TableStyleMedium9;

            // Add a new worksheet where the table will be duplicated
            Worksheet destSheet = workbook.Worksheets.Add("Copy");

            // Define source range (including header) – same as the table range
            AsposeRange srcRange = sourceSheet.Cells.CreateRange("A1:B4");

            // Determine size of the source range
            int rows = srcRange.RowCount;
            int cols = srcRange.ColumnCount;

            // Create a destination range of the same size starting at D1 (row 0, column 3)
            AsposeRange destRange = destSheet.Cells.CreateRange(0, 3, rows, cols);

            // Configure paste options to keep table objects during the copy
            PasteOptions pasteOptions = new PasteOptions
            {
                KeepOldTables = true // preserve table formatting and formulas
            };

            // Copy the source table range to the destination range with the specified options
            destRange.Copy(srcRange, pasteOptions);

            // Save the workbook containing the duplicated table
            string outputPath = "DuplicateTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
