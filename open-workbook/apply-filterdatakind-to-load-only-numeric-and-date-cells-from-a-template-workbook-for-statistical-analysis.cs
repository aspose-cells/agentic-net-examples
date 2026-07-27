// Title: C# – Load numeric and date cells with Aspose.Cells LoadFilter for statistical analysis
// Description: Shows how to use Aspose.Cells LoadOptions together with a LoadFilter (LoadDataFilterOptions.CellNumeric) to open an Excel workbook, read only numeric and date cells, calculate their sum, and optionally save the filtered result. This approach reduces memory usage and speeds up statistical calculations.
// Keywords: Aspose.Cells LoadFilter | LoadDataFilterOptions.CellNumeric | C# load numeric cells | Excel date cells as numbers | statistical analysis Aspose.Cells | sum numeric cells C# | filter workbook cells by type | memory‑efficient Excel loading
// Common Searches: load only numeric cells Aspose.Cells .NET | Aspose.Cells filter out text cells | calculate sum of numeric cells using Aspose.Cells | LoadOptions CellNumeric example | how to read dates as numbers with Aspose.Cells
// Developer Intent: Open an Excel file while loading only numeric and date cells to perform fast statistical calculations.
// Use Cases: Compute totals or averages from large financial sheets without loading text or formulas, saving memory. | Generate aggregated metrics from big data sets by loading only numeric values for quick analysis. | Create a lightweight copy of a workbook that contains only numeric data for downstream reporting or machine‑learning pipelines.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with LoadFilter to load only numeric and date cells and calculate the average. | Explain the behavior of LoadDataFilterOptions.CellNumeric in Aspose.Cells and why Excel dates are treated as numeric values. | Show how to extend the example to also load boolean cells while still summing numeric values.

using System;
using Aspose.Cells;

// Shows how to use Aspose.Cells LoadOptions together with a LoadFilter (LoadDataFilterOptions.CellNumeric) to open an Excel workbook, read only numeric and date cells, calculate their sum, and optionally save the filtered result. This approach reduces memory usage and speeds up statistical calculations.
class LoadNumericAndDateCells
{
    static void Main()
    {
        // Create LoadOptions and assign a LoadFilter that loads only numeric (including dates) cells
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellNumeric);

        // Load the workbook from a template file using the specified load options
        Workbook workbook = new Workbook("Template.xlsx", loadOptions);

        // Perform a simple statistical analysis: calculate the sum of all numeric/date cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        double sum = 0;
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                // CellValueType.IsNumeric covers both pure numbers and dates
                if (cell.Type == CellValueType.IsNumeric)
                {
                    sum += cell.DoubleValue;
                }
            }
        }

        Console.WriteLine($"Sum of numeric/date cells: {sum}");

        // Save the workbook after analysis (optional)
        workbook.Save("FilteredOutput.xlsx");
    }
}
