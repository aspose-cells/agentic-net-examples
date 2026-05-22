using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Create source table ----------
            // Header in A1
            worksheet.Cells["A1"].PutValue("Item");
            // Sample items in A2:A5
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["A5"].PutValue("Date");

            // Add a ListObject (Excel table) that covers A1:A5
            int tableIdx = worksheet.ListObjects.Add("A1", "A5", true);
            ListObject table = worksheet.ListObjects[tableIdx];
            table.DisplayName = "ItemsTable"; // Optional friendly name

            // ---------- Create data validation ----------
            // Validation will be placed in cell B1
            Validation validation = worksheet.Cells["B1"].GetValidation();

            // Restrict entries to a list type
            validation.Type = ValidationType.List;

            // Use a structured reference to the column of the table as the source list
            validation.Formula1 = "=ItemsTable[Item]";

            // Show a drop‑down arrow in the cell
            validation.InCellDropDown = true;

            // Explicitly apply the validation to B1 (optional)
            CellArea area = CellArea.CreateCellArea(0, 1, 0, 1); // Row 0, Column 1 => B1
            validation.AddArea(area);

            // ---------- Save the workbook ----------
            workbook.Save("TableValidationDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}