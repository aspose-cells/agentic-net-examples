// Title: Load Only Visible Worksheets with LightCells in Aspose.Cells for .NET
// Description: Demonstrates how to use a custom LoadFilter to load full data only for worksheets where IsVisible is true, while hidden sheets are loaded as structure only. The LightCellsDataHandler processes each visible sheet sequentially, allowing row and cell operations before saving the workbook. This approach reduces memory usage and speeds up processing when hidden worksheets are irrelevant.
// Keywords: Aspose.Cells LoadFilter visible sheets | LightCells visible worksheets .NET | C# load only visible worksheets | skip hidden worksheets Aspose.Cells | LightCellsDataHandler example | memory‑efficient Excel processing | Aspose.Cells workbook filtering
// Common Searches: Aspose.Cells load only visible worksheets C# | LightCells process visible sheets only | How to skip hidden worksheets with LoadFilter | C# example for LightCellsDataHandler visible sheets | Reduce memory usage when loading Excel with Aspose.Cells
// Developer Intent: Load a workbook so that only visible worksheets are fully loaded, then process those sheets sequentially using LightCells.
// Use Cases: Extract or transform data from visible tabs while ignoring hidden ones. | Add markers, formulas, or formatting to visible sheets after LightCells processing. | Generate reports that include only user‑visible worksheets, improving performance and memory consumption.
// AI Prompts: Write C# code that uses Aspose.Cells LoadFilter to load only visible worksheets and processes them with LightCellsDataHandler. | Explain how to modify VisibleSheetHandler to skip rows based on a custom condition while still handling only visible sheets. | Show how to combine VisibleSheetLoadFilter with column‑level filtering for a LightCells operation.

using System;
using Aspose.Cells;

namespace AsposeCellsVisibleSheetsLightCells
{
    // Custom LoadFilter that loads only visible worksheets.
    // Demonstrates how to use a custom LoadFilter to load full data only for worksheets where IsVisible is true, while hidden sheets are loaded as structure only. The LightCellsDataHandler processes each visible sheet sequentially, allowing row and cell operations before saving the workbook. This approach reduces memory usage and speeds up processing when hidden worksheets are irrelevant.
    class VisibleSheetLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data for visible sheets, only structure for hidden ones.
            if (sheet.IsVisible)
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            else
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }

    // LightCellsDataHandler that processes only visible worksheets.
    class VisibleSheetHandler : LightCellsDataHandler
    {
        // Called before reading a worksheet. Return true only for visible sheets.
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Start processing sheet: {sheet.Name} (Visible={sheet.IsVisible})");
            return sheet.IsVisible;
        }

        // Process each row – return true to read its cells.
        public bool StartRow(int rowIndex)
        {
            // All rows in a visible sheet are processed.
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No special row processing needed.
            return true;
        }

        // Process each cell – return true to read the cell.
        public bool StartCell(int columnIndex)
        {
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // Example processing: output cell address and value.
            Console.WriteLine($"  Cell {cell.Name}: {cell.Value}");
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook.
            string inputPath = "input.xlsx";
            // Path for the resulting workbook (optional, can be same as input).
            string outputPath = "output.xlsx";

            // Configure load options with the custom filter and handler.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new VisibleSheetLoadFilter(),
                LightCellsDataHandler = new VisibleSheetHandler()
            };

            // Load the workbook using LightCells mode.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // At this point only visible worksheets are loaded.
            Console.WriteLine($"Total worksheets loaded: {workbook.Worksheets.Count}");

            // Optional: further processing after load can be done here.
            // For demonstration, iterate through the loaded worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Processing after load: {sheet.Name}");
                // Example: write a marker in A1 of each visible sheet.
                sheet.Cells["A1"].PutValue($"Processed {DateTime.Now}");
            }

            // Save the workbook (preserves only the visible sheets).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
