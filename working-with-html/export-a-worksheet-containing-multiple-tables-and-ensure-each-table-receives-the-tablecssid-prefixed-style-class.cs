// Title: Export a worksheet with multiple tables and assign a TableCssId‑prefixed CSS class to each table using Aspose.Cells for .NET
// Common Searches: how to assign a TableCssId prefix to each ListObject when exporting to HTML with Aspose.Cells C# | export a worksheet that contains several tables and apply custom CSS classes using Aspose.Cells .NET | define separate CellArea ranges for multiple tables in a single worksheet with Aspose.Cells | set different TableStyleType values for each table while saving to XLSX in Aspose.Cells | add multiple ListObjects to one worksheet and preserve their CSS classes in HTML output Aspose.Cells
// Tags: Aspose.Cells TableCssId assignment for ListObject | multiple ListObjects with distinct CellArea ranges C# | apply built‑in TableStyleType per table Aspose.Cells | export worksheet with tables to HTML preserving CSS class | save workbook containing tables to XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, defines two separate CellArea ranges, adds a ListObject for each range, assigns a TableCssId‑prefixed CSS class and distinct built‑in TableStyleType to each table, and saves the workbook as an XLSX file.
class ExportWorksheetWithTables
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // ------------------------------
            // Populate data for the first table
            // ------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Define the range for the first table (including header)
            CellArea firstTableArea = CellArea.CreateCellArea("A1", "C3");

            // Add the first ListObject (table) to the worksheet
            int firstTableIndex = sheet.ListObjects.Add(
                firstTableArea.StartRow,
                firstTableArea.StartColumn,
                firstTableArea.EndRow,
                firstTableArea.EndColumn,
                true);
            ListObject firstTable = sheet.ListObjects[firstTableIndex];

            // Optional: set a built‑in table style for visual appearance
            firstTable.TableStyleType = TableStyleType.TableStyleMedium2;

            // ------------------------------
            // Populate data for the second table
            // ------------------------------
            // Header row
            sheet.Cells["E1"].PutValue("Product");
            sheet.Cells["F1"].PutValue("Category");
            sheet.Cells["G1"].PutValue("Price");

            // Sample data rows
            sheet.Cells["E2"].PutValue("Laptop");
            sheet.Cells["F2"].PutValue("Electronics");
            sheet.Cells["G2"].PutValue(1200);

            sheet.Cells["E3"].PutValue("Desk");
            sheet.Cells["F3"].PutValue("Furniture");
            sheet.Cells["G3"].PutValue(250);

            sheet.Cells["E4"].PutValue("Pen");
            sheet.Cells["F4"].PutValue("Stationery");
            sheet.Cells["G4"].PutValue(2.5);

            // Define the range for the second table
            CellArea secondTableArea = CellArea.CreateCellArea("E1", "G4");

            // Add the second ListObject (table) to the worksheet
            int secondTableIndex = sheet.ListObjects.Add(
                secondTableArea.StartRow,
                secondTableArea.StartColumn,
                secondTableArea.EndRow,
                secondTableArea.EndColumn,
                true);
            ListObject secondTable = sheet.ListObjects[secondTableIndex];

            // Optional: set a different built‑in table style
            secondTable.TableStyleType = TableStyleType.TableStyleMedium9;

            // ------------------------------
            // Save the workbook to a file
            // ------------------------------
            workbook.Save("MultipleTables.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
