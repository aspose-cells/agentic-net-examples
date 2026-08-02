using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the table (column A)
        worksheet.Cells["A1"].PutValue("Item");   // Header
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["A5"].PutValue("Date");

        // Add a ListObject (Excel table) that covers the range A1:A5
        int tableIndex = worksheet.ListObjects.Add("A1", "A5", true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "ItemsTable"; // Optional: give the table a friendly name

        // Define the cell (B1) where the data validation will be applied
        CellArea validationArea = new CellArea
        {
            StartRow = 0,      // Row 0 (B1)
            StartColumn = 1,   // Column B
            EndRow = 0,
            EndColumn = 1
        };

        // Add a validation to the worksheet for the specified cell area
        int validationIdx = worksheet.Validations.Add(validationArea);
        Validation validation = worksheet.Validations[validationIdx];

        // Configure the validation as a List that pulls values from the table column
        validation.Type = ValidationType.List;
        // Reference the table column using the table name and column header
        validation.Formula1 = "ItemsTable[Item]";
        validation.InCellDropDown = true; // Show the dropdown arrow in the cell

        // Save the workbook
        workbook.Save("DataValidationFromTable.xlsx");
    }
}