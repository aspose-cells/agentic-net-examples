// Title: Create a list data validation that automatically expands when inserting rows using CellArea in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a list validation for column A and automatically extends its range after a new row is inserted with InsertRange. | Demonstrate how to combine CellArea.CreateCellArea with Validation.AddArea to keep a dropdown list validation dynamic as rows are added. | Provide a complete example that saves the workbook after expanding the validation to include the newly inserted row.
// Common Searches: aspocells keep dropdown validation range updated after inserting rows c# | dynamic list validation using CellArea in Aspose.Cells example | c# expand data validation when adding new rows with InsertRange | how to extend validation area after row insertion Aspose.Cells
// Tags: Aspose.Cells dynamic list validation | CellArea validation range expansion | InsertRange shift down update validation | C# data validation dropdown growth

using Aspose.Cells;

// The example creates a list validation for cells A2:A5, inserts a new row at position 3, expands the validation to include the new row using Validation.AddArea, and saves the workbook as an XLSX file.
class DynamicValidationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Fill column A with sample data (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Item {i + 1}");
        }

        // Create a validation that initially covers A2:A5 (rows 1‑4, column 0)
        CellArea initialArea = CellArea.CreateCellArea(1, 0, 4, 0);
        int validationIndex = ws.Validations.Add(initialArea);
        Validation validation = ws.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.InCellDropDown = true;
        validation.Formula1 = "\"Option1,Option2,Option3\"";

        // Insert a new row at position 3 (zero‑based row index 2) and shift cells down
        CellArea insertArea = CellArea.CreateCellArea(2, 0, 2, 0);
        cells.InsertRange(insertArea, ShiftType.Down);

        // Expand the validation to include the newly inserted row (row 3, column A)
        CellArea newRowArea = CellArea.CreateCellArea(2, 0, 2, 0);
        validation.AddArea(newRowArea, false, false);

        // Save the workbook
        wb.Save("DynamicValidationDemo.xlsx", SaveFormat.Xlsx);
    }
}
