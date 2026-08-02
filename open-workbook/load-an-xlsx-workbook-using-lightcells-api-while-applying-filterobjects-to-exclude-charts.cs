// Title: Load an XLSX workbook with LightCells API while excluding charts – Aspose.Cells for .NET
// Description: Demonstrates how to load an XLSX file in LightCells mode using a custom LoadFilter that clears the Chart flag, paired with a simple LightCellsDataHandler. The workbook is loaded without any chart objects, reducing memory usage and speeding up processing. The sample also shows how to verify that worksheets are present and charts are omitted.
// Keywords: Aspose.Cells LightCells load workbook | C# LoadFilter exclude charts | LightCellsDataHandler example | skip chart objects Aspose.Cells | memory‑efficient Excel loading .NET | ChartExcludingLoadFilter | load XLSX without charts
// Common Searches: Aspose.Cells LightCells load workbook without charts | C# filter to skip charts when loading Excel | How to use LoadFilter to exclude chart objects | LightCellsDataHandler sample for chart exclusion | Load large XLSX files quickly with Aspose.Cells
// Developer Intent: Load an XLSX workbook in LightCells mode while preventing chart objects from being loaded into memory.
// Use Cases: Fast, low‑memory loading of large spreadsheets when only cell data is needed. | Processing worksheets, rows, and cells with a LightCellsDataHandler without the overhead of chart objects. | Generating data‑only reports or performing analytics where charts are irrelevant.
// AI Prompts: Show a C# example that uses Aspose.Cells LightCells API to load an XLSX file and exclude charts with a custom LoadFilter. | Explain how to modify the LoadFilter to also skip pictures, shapes, or other drawing objects while using LightCells. | Provide code to count charts after loading a workbook to confirm they were omitted.

using System;
using Aspose.Cells;

// Demonstrates how to load an XLSX file in LightCells mode using a custom LoadFilter that clears the Chart flag, paired with a simple LightCellsDataHandler. The workbook is loaded without any chart objects, reducing memory usage and speeding up processing. The sample also shows how to verify that worksheets are present and charts are omitted.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string filePath = "input.xlsx";

        // Create a LoadFilter that excludes charts from being loaded
        LoadFilter filter = new ChartExcludingLoadFilter();

        // Create a LightCellsDataHandler that simply accepts all data
        LightCellsDataHandler handler = new SimpleLightCellsHandler();

        // Configure load options with the filter and the LightCells handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = filter;
        loadOptions.LightCellsDataHandler = handler;

        // Load the workbook using LightCells mode; charts will be omitted
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Demonstrate that worksheets are loaded and charts are excluded
        Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);
        int totalCharts = 0;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            totalCharts += ws.Charts.Count;
        }
        Console.WriteLine("Charts loaded (should be 0): " + totalCharts);
    }

    // Custom LoadFilter that disables chart loading for every sheet
    class ChartExcludingLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load all data except charts
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        }
    }

    // Simple LightCellsDataHandler that processes all sheets, rows, and cells
    class SimpleLightCellsHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet) => true;          // Process every sheet
        public bool StartRow(int rowIndex) => true;              // Process every row
        public bool ProcessRow(Row row) => true;                 // No custom row processing
        public bool StartCell(int columnIndex) => true;          // Process every cell
        public bool ProcessCell(Cell cell) => true;              // No custom cell processing
    }
}
