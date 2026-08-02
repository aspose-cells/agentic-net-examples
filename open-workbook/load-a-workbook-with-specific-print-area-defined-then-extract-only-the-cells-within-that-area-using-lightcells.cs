// Title: C# – Extract Cells from a Defined Print Area Using Aspose.Cells LightCells
// Description: This example shows how to load an Excel workbook in LightCells mode with a custom LightCellsDataHandler that reads the worksheet's print area (or the full used range when no print area exists) and captures only the cells inside that region. The handler stores each cell's address and value, enabling memory‑efficient processing of large sheets.
// Keywords: Aspose.Cells LightCells C# | print area extraction | custom LightCellsDataHandler | load workbook partial range | Excel print area API | memory efficient Excel processing | CellArea CreateCellArea | Aspose.Cells GitHub example | filter cells by print area | C# Excel cell extraction
// Common Searches: Aspose.Cells load only print area | C# LightCells handler for print area | extract cells from defined print range Aspose | how to use CellArea with LightCells | partial workbook loading Aspose.Cells | sample code LightCellsDataHandler print area
// Developer Intent: Load a workbook and retrieve only the cells that belong to the worksheet's defined print area using LightCells.
// Use Cases: Process massive spreadsheets while limiting memory consumption by loading just the printable region. | Create a printable report that includes only the cells marked for printing, ignoring hidden or auxiliary data. | Validate or audit values within the print area before sending the file to a printer or export routine.
// AI Prompts: Generate a LightCellsDataHandler that records cell formulas as well as values for the defined print area. | Explain how to modify the handler to fall back to a named range when a print area is not set. | Provide a step‑by‑step guide to integrate this LightCells approach into an ASP.NET Core web API.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLightCellsPrintAreaDemo
{
    // Custom LightCellsDataHandler that extracts only cells inside the worksheet's print area
    // This example shows how to load an Excel workbook in LightCells mode with a custom LightCellsDataHandler that reads the worksheet's print area (or the full used range when no print area exists) and captures only the cells inside that region. The handler stores each cell's address and value, enabling memory‑efficient processing of large sheets.
    public class PrintAreaLightCellsHandler : LightCellsDataHandler
    {
        private CellArea _printArea;
        private readonly List<string> _extractedValues = new List<string>();

        // Expose extracted values after loading
        public IReadOnlyList<string> Values => _extractedValues.AsReadOnly();

        // Called when a worksheet starts processing
        public bool StartSheet(Worksheet sheet)
        {
            // Retrieve the print area defined in the worksheet (e.g., "A1:B3")
            string area = sheet.PageSetup.PrintArea;

            if (string.IsNullOrEmpty(area))
            {
                // If no print area is defined, consider the whole used range
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;
                _printArea = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
            }
            else
            {
                // Convert the address string to a CellArea
                string[] parts = area.Split(':');
                _printArea = CellArea.CreateCellArea(parts[0], parts[1]);
            }

            // Continue processing this sheet
            return true;
        }

        // Called for each row; return true only if the row is inside the print area
        public bool StartRow(int rowIndex)
        {
            return rowIndex >= _printArea.StartRow && rowIndex <= _printArea.EndRow;
        }

        // Called for each cell; return true only if the column is inside the print area
        public bool StartCell(int columnIndex)
        {
            return columnIndex >= _printArea.StartColumn && columnIndex <= _printArea.EndColumn;
        }

        // Process the cell data (store its address and value)
        public bool ProcessCell(Cell cell)
        {
            _extractedValues.Add($"{cell.Name}: {cell.Value}");
            return true; // Continue processing
        }

        // Row processing is not needed for this demo, but must be implemented
        public bool ProcessRow(Row row)
        {
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (must contain a defined print area)
            string inputPath = "input.xlsx";

            // Set up LoadOptions with the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions();
            PrintAreaLightCellsHandler handler = new PrintAreaLightCellsHandler();
            loadOptions.LightCellsDataHandler = handler;

            // Load the workbook using LightCells mode
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // After loading, the handler contains only the cells that belong to the print area
            Console.WriteLine("Cells extracted from the defined print area:");
            foreach (string entry in handler.Values)
            {
                Console.WriteLine(entry);
            }

            // (Optional) Save the workbook if further processing is required
            // workbook.Save("output.xlsx");
        }
    }
}
