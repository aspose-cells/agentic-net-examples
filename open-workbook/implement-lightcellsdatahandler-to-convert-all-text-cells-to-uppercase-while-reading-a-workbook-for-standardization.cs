using System;
using Aspose.Cells;

public class UpperCaseHandler : LightCellsDataHandler
{
    // Process every sheet
    public bool StartSheet(Worksheet sheet)
    {
        return true; // continue processing this sheet
    }

    // Process every row
    public bool StartRow(int rowIndex)
    {
        return true; // continue processing this row
    }

    // No special row handling needed
    public bool ProcessRow(Row row)
    {
        return true; // continue processing cells in this row
    }

    // Process every cell in the row
    public bool StartCell(int columnIndex)
    {
        return true; // continue processing this cell
    }

    // Convert text cells to uppercase
    public bool ProcessCell(Cell cell)
    {
        // Check if the cell contains a string value
        if (cell.Type == CellValueType.IsString)
        {
            // Replace the cell value with its uppercase representation
            cell.PutValue(cell.StringValue.ToUpper());
        }
        return true; // keep the cell in the workbook model
    }
}

public class Program
{
    public static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Configure load options to use the custom LightCellsDataHandler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new UpperCaseHandler();

        // Load the workbook in light cells mode; the handler will process each cell
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Save the processed workbook
        workbook.Save("output.xlsx");
    }
}