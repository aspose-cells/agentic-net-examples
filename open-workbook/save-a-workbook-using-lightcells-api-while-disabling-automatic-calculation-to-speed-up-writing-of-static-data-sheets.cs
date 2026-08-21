// Title: Save a workbook with LightCells API and disable formula calculation in Aspose.Cells for .NET
// Description: Shows how to create a Workbook, populate 1,000 static rows, set FormulaSettings.CalculateOnSave = false, implement a SimpleLightCellsDataProvider that streams worksheet data, configure OoxmlSaveOptions for LightCells mode, and save the file as an XLSX with optimal performance.
// Keywords: Aspose.Cells | LightCells API | C# | .NET | FormulaSettings.CalculateOnSave | OoxmlSaveOptions | custom LightCellsDataProvider | disable calculation | fast Excel save | performance optimization
// Common Searches: Aspose.Cells LightCells save example C# | disable formula calculation when saving Excel with Aspose.Cells | how to use SimpleLightCellsDataProvider | improve save speed using LightCells mode | stream rows to XLSX with Aspose.Cells
// Developer Intent: Generate an XLSX file quickly by using LightCells streaming while turning off automatic formula recalculation.
// Use Cases: Export large static datasets to XLSX with minimal processing time. | Create reports where formulas are unnecessary and fast write speed is critical. | Implement a custom LightCellsDataProvider to write selected sheets without triggering calculations.
// AI Prompts: Provide a C# snippet that saves a workbook with LightCells API and sets FormulaSettings.CalculateOnSave to false. | Explain how to extend SimpleLightCellsDataProvider to handle multiple worksheets while keeping calculation disabled. | Suggest additional techniques to further boost save performance when using LightCells in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to create a Workbook, populate 1,000 static rows, set FormulaSettings.CalculateOnSave = false, implement a SimpleLightCellsDataProvider that streams worksheet data, configure OoxmlSaveOptions for LightCells mode, and save the file as an XLSX with optimal performance.
class LightCellsSaveDemo
{
    static void Main()
    {
        // Create a new workbook (rule: Workbook())
        Workbook workbook = new Workbook();

        // Populate static data into the first worksheet
        Worksheet ws = workbook.Worksheets[0];
        for (int i = 0; i < 1000; i++)
        {
            ws.Cells[i, 0].PutValue($"Row {i}");
            ws.Cells[i, 1].PutValue(i * 10);
        }

        // Disable automatic formula calculation on save to improve performance
        // (rule: FormulaSettings.CalculateOnSave property)
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Create a LightCells data provider that streams the worksheet data
        LightCellsDataProvider provider = new SimpleLightCellsDataProvider(ws);

        // Configure save options to use LightCells mode (rule: OoxmlSaveOptions.LightCellsDataProvider)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = provider
        };

        // Save the workbook using LightCells API (rule: Workbook.Save(string, SaveOptions))
        workbook.Save("LightCellsOutput.xlsx", saveOptions);
    }
}

// Simple implementation of LightCellsDataProvider for a single worksheet
class SimpleLightCellsDataProvider : LightCellsDataProvider
{
    private readonly Worksheet _worksheet;
    private int _currentRow = -1;
    private int _currentCol = -1;
    private readonly int _maxRow;
    private readonly int _maxCol;

    public SimpleLightCellsDataProvider(Worksheet worksheet)
    {
        _worksheet = worksheet;
        // Determine the used range of the worksheet
        _maxRow = worksheet.Cells.MaxDataRow;
        _maxCol = worksheet.Cells.MaxDataColumn;
    }

    // Process only the first sheet (index 0)
    public bool StartSheet(int sheetIndex) => sheetIndex == 0;

    // Return the next row index or -1 when done
    public int NextRow()
    {
        if (_currentRow < _maxRow)
        {
            _currentRow++;
            _currentCol = -1;
            return _currentRow;
        }
        return -1;
    }

    // No special row initialization needed
    public void StartRow(Row row) { }

    // Return the next cell (column) index or -1 when done
    public int NextCell()
    {
        if (_currentCol < _maxCol)
        {
            _currentCol++;
            return _currentCol;
        }
        return -1;
    }

    // Copy the value from the source worksheet cell to the target cell
    public void StartCell(Cell cell)
    {
        Cell src = _worksheet.Cells[_currentRow, _currentCol];
        if (src.Type != CellValueType.IsNull)
        {
            cell.PutValue(src.Value);
        }
    }

    // Allow Aspose to gather strings into the global pool
    public bool IsGatherString() => true;
}
