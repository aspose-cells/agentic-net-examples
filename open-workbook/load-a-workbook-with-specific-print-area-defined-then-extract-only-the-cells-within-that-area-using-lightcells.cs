using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLightCellsExample
{
    // Custom LightCellsDataHandler that extracts cells inside a given print area
    public class PrintAreaExtractor : LightCellsDataHandler
    {
        private readonly CellArea _printArea;
        public List<string> ExtractedValues { get; } = new List<string>();

        public PrintAreaExtractor(CellArea printArea)
        {
            _printArea = printArea;
        }

        // Process each worksheet – continue processing all sheets
        public bool StartSheet(Worksheet sheet) => true;

        // Process each row – continue processing all rows
        public bool StartRow(int rowIndex) => true;

        // Process row data – not needed for extraction, just continue
        public bool ProcessRow(Row row) => true;

        // Process each cell – continue processing all cells
        public bool StartCell(int columnIndex) => true;

        // Extract cell value if it lies within the print area
        public bool ProcessCell(Cell cell)
        {
            if (cell.Row >= _printArea.StartRow && cell.Row <= _printArea.EndRow &&
                cell.Column >= _printArea.StartColumn && cell.Column <= _printArea.EndColumn)
            {
                // Store cell address and its value
                ExtractedValues.Add($"{cell.Name}: {cell.Value}");
            }
            return true; // Continue processing
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (must have a print area defined)
            string sourceFile = "input.xlsx";

            // -----------------------------------------------------------------
            // Step 1: Load the workbook normally to read the defined print area
            // -----------------------------------------------------------------
            Workbook tempWb = new Workbook(sourceFile);
            Worksheet tempWs = tempWb.Worksheets[0];

            // Retrieve the print area string (e.g., "A1:C5")
            string printAreaStr = tempWs.PageSetup.PrintArea;
            if (string.IsNullOrEmpty(printAreaStr))
            {
                Console.WriteLine("No print area defined in the worksheet.");
                return;
            }

            // Convert the print area string to a CellArea object
            // Assuming the format is a single range like "A1:C5"
            string[] parts = printAreaStr.Split(':');
            CellArea printArea = CellArea.CreateCellArea(parts[0], parts[1]);

            // ---------------------------------------------------------------
            // Step 2: Load the workbook using LightCells with a custom handler
            // ---------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                // Attach the custom handler that will collect cells inside the print area
                LightCellsDataHandler = new PrintAreaExtractor(printArea)
            };

            // Load the workbook in LightCells mode
            Workbook wb = new Workbook(sourceFile, loadOptions);

            // Retrieve the handler to access extracted data
            PrintAreaExtractor handler = (PrintAreaExtractor)loadOptions.LightCellsDataHandler;

            // -------------------------------------------------
            // Step 3: Output the extracted cells from the area
            // -------------------------------------------------
            Console.WriteLine($"Extracted cells from print area \"{printAreaStr}\":");
            foreach (string entry in handler.ExtractedValues)
            {
                Console.WriteLine(entry);
            }

            // (Optional) Save the workbook if further processing is needed
            // wb.Save("output.xlsx");
        }
    }
}