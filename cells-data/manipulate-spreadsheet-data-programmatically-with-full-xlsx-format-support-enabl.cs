using System;
using Aspose.Cells;

class SpreadsheetEditor
{
    static void Main()
    {
        // Load an existing XLSX file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Update a specific cell value
        cells["B2"].PutValue(12345);

        // Insert a new column at index 2 (C column)
        sheet.Cells.InsertColumn(2);

        // Set header for the new column
        cells["C1"].PutValue("NewColumn");

        // Populate the new column based on values from column A
        int lastRow = cells.MaxDataRow;
        for (int row = 1; row <= lastRow; row++)
        {
            Cell sourceCell = cells[row, 0]; // column A (0‑based index)
            if (sourceCell.Type == CellValueType.IsNumeric)
            {
                double sourceValue = sourceCell.DoubleValue;
                cells[row, 2].PutValue(sourceValue * 2); // column C (new column)
            }
            else
            {
                // If not numeric, copy the original value as text
                cells[row, 2].PutValue(sourceCell.StringValue);
            }
        }

        // Recalculate all formulas after modifications
        workbook.CalculateFormula();

        // Replace any placeholder text "PLACEHOLDER" with a real value
        workbook.Replace("PLACEHOLDER", "ActualValue");

        // Save the edited workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}