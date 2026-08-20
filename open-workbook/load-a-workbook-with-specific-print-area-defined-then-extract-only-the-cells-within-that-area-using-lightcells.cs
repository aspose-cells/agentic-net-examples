// Title: Aspose.Cells .NET – Extract Cells from a Defined Print Area Using LightCells
// Description: Loads an Excel workbook in LightCells mode with a custom LightCellsDataHandler that reads each worksheet's PageSetup.PrintArea, converts it to a CellArea, and streams only the cells inside that range. Extracted cell addresses and values are stored in a dictionary for fast, memory‑efficient processing.
// Keywords: Aspose.Cells LightCells | print area extraction .NET | LoadOptions LightCellsDataHandler | CellArea filter rows columns | stream Excel cells efficiently | custom LightCells handler example | Aspose.Cells C# print area
// Common Searches: How to read only the print area with Aspose.Cells LightCells | Extract cell values from a specific print range in .NET | LightCellsDataHandler example for print area filtering | Load workbook efficiently and limit to printable region | Aspose.Cells C# extract cells by PageSetup.PrintArea
// Developer Intent: Load a workbook in LightCells mode and retrieve only the cells that belong to the worksheet's defined print area.
// Use Cases: Generate a report that includes just the printable section of a template workbook. | Copy or export data from the print area to another file or database without loading the full sheet. | Validate content inside the print area during automated QA tests.
// AI Prompts: Create a LightCellsDataHandler that returns a list of cell addresses and their values for the defined print area. | Extend the PrintAreaHandler to also capture cell formulas and formatting information. | Show how to process multiple worksheets, each with its own print area, using LightCells in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook in LightCells mode with a custom LightCellsDataHandler that reads each worksheet's PageSetup.PrintArea, converts it to a CellArea, and streams only the cells inside that range. Extracted cell addresses and values are stored in a dictionary for fast, memory‑efficient processing.
class PrintAreaExtractor
{
    static void Main()
    {
        // Path to the workbook that has a print area defined
        string filePath = "input.xlsx";

        // Create a custom LightCellsDataHandler that extracts cells inside the print area
        var handler = new PrintAreaHandler();

        // Configure load options to use the handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = handler;

        // Load the workbook in LightCells mode
        Workbook workbook = new Workbook(filePath, loadOptions);

        // After loading, the handler contains the extracted cells
        Console.WriteLine("Cells extracted from the defined print area:");
        foreach (var kvp in handler.ExtractedData)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }
    }
}

// Custom handler implementing LightCellsDataHandler
public class PrintAreaHandler : LightCellsDataHandler
{
    // Current worksheet print area
    private CellArea _printArea;

    // Collected cell address/value pairs
    public Dictionary<string, object> ExtractedData { get; } = new Dictionary<string, object>();

    // Called when a new worksheet is encountered
    public bool StartSheet(Worksheet sheet)
    {
        // Retrieve the print area string (e.g., "A1:B3")
        string areaStr = sheet.PageSetup.PrintArea;

        if (!string.IsNullOrEmpty(areaStr))
        {
            // Convert the string to a CellArea object
            var parts = areaStr.Split(':');
            if (parts.Length == 2)
                _printArea = CellArea.CreateCellArea(parts[0], parts[1]);
            else
                _printArea = new CellArea(); // fallback to empty area
        }
        else
        {
            // No print area defined – set an impossible range so nothing is processed
            _printArea = new CellArea { StartRow = -1 };
        }

        // Continue processing rows in this sheet
        return true;
    }

    // Called before processing each row
    public bool StartRow(int rowIndex)
    {
        // Process the row only if its index lies within the print area rows
        return rowIndex >= _printArea.StartRow && rowIndex <= _printArea.EndRow;
    }

    // Called after a row is started; return true to continue to its cells
    public bool ProcessRow(Row row)
    {
        return true;
    }

    // Called before processing each cell in the current row
    public bool StartCell(int columnIndex)
    {
        // Process the cell only if its column lies within the print area columns
        return columnIndex >= _printArea.StartColumn && columnIndex <= _printArea.EndColumn;
    }

    // Called for each cell that passed the StartCell check
    public bool ProcessCell(Cell cell)
    {
        // Store the cell's address (e.g., "B2") and its value
        ExtractedData[cell.Name] = cell.Value;
        return true;
    }
}
