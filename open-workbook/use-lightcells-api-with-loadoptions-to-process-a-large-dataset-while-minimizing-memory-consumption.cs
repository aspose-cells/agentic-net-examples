using System;
using Aspose.Cells;

public class LargeDatasetHandler : LightCellsDataHandler
{
    // Called once for each worksheet that is being read.
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Start processing sheet: {sheet.Name}");
        // Return true to continue processing this sheet.
        return true;
    }

    // Called before a row is read.
    public bool StartRow(int rowIndex)
    {
        // Return true to read this row.
        return true;
    }

    // Called after a row has been read.
    public bool ProcessRow(Row row)
    {
        // No special row processing needed; just continue.
        return true;
    }

    // Called before a cell in the current row is read.
    public bool StartCell(int columnIndex)
    {
        // Return true to read this cell.
        return true;
    }

    // Called after a cell has been read.
    public bool ProcessCell(Cell cell)
    {
        // Example: output cell address and its value.
        Console.WriteLine($"Cell[{cell.Row},{cell.Column}] = {cell.Value}");
        return true;
    }
}

public class Program
{
    public static void Main()
    {
        // Create LoadOptions and attach the LightCellsDataHandler.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new LargeDatasetHandler();

        // Disable keeping unparsed data to reduce memory usage.
        loadOptions.KeepUnparsedData = false;

        // Load the large workbook in streaming (LightCells) mode.
        Workbook workbook = new Workbook("LargeData.xlsx", loadOptions);

        // Set the workbook's cells to use MemoryPreference for compact storage.
        workbook.Worksheets[0].Cells.MemorySetting = MemorySetting.MemoryPreference;

        // Save the workbook after processing (can be the same or a new file).
        workbook.Save("ProcessedLargeData.xlsx");
    }
}