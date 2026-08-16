// Title: C# – Load an XLSX workbook with Aspose.Cells LightCells while excluding charts
// Description: Shows how to open an XLSX file in LightCells mode using Aspose.Cells for C# and prevent chart objects from being loaded. The example builds a LoadDataFilterOptions value that omits the Chart flag, applies it via a LoadFilter, attaches a minimal LightCellsDataHandler, and then loads the workbook. Worksheet and chart counts are printed to verify that charts are excluded, and the workbook can be saved afterward.
// Keywords: Aspose.Cells LightCells C# | LoadDataFilterOptions exclude charts | LoadFilter LoadOptions example | LightCellsDataHandler minimal implementation | load XLSX without charts | memory‑efficient workbook loading | chart omission Aspose.Cells | C# Excel processing LightCells
// Common Searches: Aspose.Cells LightCells load workbook without charts C# | How to skip chart objects when loading XLSX with Aspose.Cells | LoadDataFilterOptions chart flag C# example | Minimal LightCellsDataHandler code sample | Exclude charts using LoadFilter in Aspose.Cells
// Developer Intent: Load an XLSX workbook in LightCells mode with Aspose.Cells for C# while filtering out all chart objects.
// Use Cases: Process large Excel files for data analysis without the overhead of chart objects. | Iterate through worksheets, rows, and cells using a custom LightCellsDataHandler while deliberately ignoring charts. | Save a modified workbook after processing, ensuring the output contains no chart data.
// AI Prompts: Generate C# code that uses Aspose.Cells LightCells to load a workbook and exclude charts by configuring LoadDataFilterOptions. | Provide a minimal LightCellsDataHandler implementation that processes sheets, rows, and cells but does not handle charts. | Explain step‑by‑step how to combine LoadFilter and LoadOptions to prevent chart loading in Aspose.Cells and verify the result.

using System;
using Aspose.Cells;

// Shows how to open an XLSX file in LightCells mode using Aspose.Cells for C# and prevent chart objects from being loaded. The example builds a LoadDataFilterOptions value that omits the Chart flag, applies it via a LoadFilter, attaches a minimal LightCellsDataHandler, and then loads the workbook. Worksheet and chart counts are printed to verify that charts are excluded, and the workbook can be saved afterward.
class Program
{
    static void Main()
    {
        // Exclude charts by removing the Chart flag from the default All options
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        LoadFilter loadFilter = new LoadFilter(filterOptions);

        // Configure load options with the filter and a LightCells data handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = loadFilter;
        loadOptions.LightCellsDataHandler = new SimpleLightCellsHandler();

        // Load the workbook using LightCells mode; charts will not be loaded
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Demonstrate that worksheets are loaded and charts are excluded
        Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Sheet '{sheet.Name}' contains {sheet.Charts.Count} chart(s).");
        }

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }

    // Minimal LightCellsDataHandler implementation required for LightCells mode
    class SimpleLightCellsHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process all rows
            return true;
        }

        public bool ProcessRow(Row row)
        {
            // No custom row processing needed
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process all cells
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // No custom cell processing needed
            return true;
        }
    }
}
