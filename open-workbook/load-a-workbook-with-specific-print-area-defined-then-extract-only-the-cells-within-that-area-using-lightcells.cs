// Title: Read Only the Print Area with LightCells in Aspose.Cells for .NET
// Description: Loads an Excel file, retrieves the worksheet's PageSetup.PrintArea, converts it to a CellArea, and uses a LightCellsDataHandler to process and output only the cells that fall inside that printable range.
// Keywords: Aspose.Cells LightCells | C# print area extraction | CellArea CreateCellArea | PageSetup.PrintArea Aspose | read selective cells .NET | memory‑efficient Excel processing
// Common Searches: Aspose.Cells LightCells read only print area | extract cells from defined print range C# | how to use PageSetup.PrintArea with LightCells | convert print area string to CellArea Aspose | filter rows by print area in Aspose.Cells
// Developer Intent: Read a workbook’s defined printable region and handle only those cells via LightCells to reduce memory usage.
// Use Cases: Generate a quick report of values inside the printable area without loading the full sheet. | Export the printable region of a massive spreadsheet to CSV or PDF while keeping memory footprint low. | Validate data confined to the print area for compliance checks in automated pipelines.
// AI Prompts: Show a C# LightCellsDataHandler that writes each extracted print‑area cell to a CSV file. | Explain how to modify PrintAreaHandler to skip empty cells and count non‑empty cells during processing. | Provide an example that iterates over multiple worksheets, each with its own print area, using LightCells in Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel file, retrieves the worksheet's PageSetup.PrintArea, converts it to a CellArea, and uses a LightCellsDataHandler to process and output only the cells that fall inside that printable range.
class PrintAreaHandler : LightCellsDataHandler
{
    private readonly CellArea _printArea;

    public PrintAreaHandler(CellArea printArea)
    {
        _printArea = printArea;
    }

    // Called for each worksheet; we process all sheets.
    public bool StartSheet(Worksheet sheet) => true;

    // Process only rows that lie inside the print area.
    public bool StartRow(int rowIndex)
    {
        return rowIndex >= _printArea.StartRow && rowIndex <= _printArea.EndRow;
    }

    // Continue processing the row once it is accepted.
    public bool ProcessRow(Row row) => true;

    // Process only columns that lie inside the print area.
    public bool StartCell(int columnIndex)
    {
        return columnIndex >= _printArea.StartColumn && columnIndex <= _printArea.EndColumn;
    }

    // Output the cell address and its value.
    public bool ProcessCell(Cell cell)
    {
        Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
        return true;
    }
}

class Program
{
    static void Main()
    {
        // Path to the workbook that already has a print area defined.
        string filePath = "input.xlsx";

        // Load the workbook normally to read the defined print area.
        Workbook tempWb = new Workbook(filePath);
        Worksheet firstSheet = tempWb.Worksheets[0];
        string printArea = firstSheet.PageSetup.PrintArea; // e.g., "A1:C5"

        if (string.IsNullOrEmpty(printArea))
        {
            Console.WriteLine("The workbook does not have a print area defined.");
            return;
        }

        // Convert the print area string to a CellArea object.
        string[] areaParts = printArea.Split(':');
        CellArea area = CellArea.CreateCellArea(areaParts[0], areaParts[1]);

        // Set up LightCellsDataHandler to extract only cells inside the print area.
        var handler = new PrintAreaHandler(area);
        LoadOptions loadOptions = new LoadOptions
        {
            LightCellsDataHandler = handler
        };

        // Load the workbook using LightCells mode; the handler will process the cells.
        Workbook lightWb = new Workbook(filePath, loadOptions);

        // No further actions required – the handler has already output the cells.
    }
}
