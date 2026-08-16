// Title: C# – Load Only Numeric & Date Cells from an Excel Template with Aspose.Cells LoadFilter for Fast Statistics
// Description: Demonstrates how to create LoadOptions with a LoadFilter (LoadDataFilterOptions.CellNumeric) to load only numeric and date cells from "Template.xlsx", compute a column sum, and save the result as "Processed.xlsx". This approach reduces memory usage and speeds up statistical calculations.
// Keywords: Aspose.Cells | LoadFilter | CellNumeric | load numeric cells C# | load date cells Aspose | Excel performance optimization | statistical analysis Excel | memory‑efficient workbook loading | C# Excel filtering
// Common Searches: Aspose.Cells load only numeric cells | LoadFilter to exclude text in Excel C# | CellNumeric option example | How to sum a column after filtering workbook cells | Improve Excel processing speed with LoadOptions
// Developer Intent: Load a workbook while filtering out non‑numeric content so that only numeric and date values are available for fast statistical calculations.
// Use Cases: Process large financial templates by loading only numbers and dates, then calculate totals without the overhead of text cells. | Perform date‑driven aggregations on massive datasets while keeping memory consumption low. | Generate lightweight reporting workbooks that contain solely the numeric data needed for charts or further analysis.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with LoadFilter to load only numeric and date cells from an Excel file and then compute the average of a chosen column. | Explain why LoadDataFilterOptions.CellNumeric treats Excel dates as numeric values and show how to confirm cell types after loading. | Provide performance‑tuning tips for using LoadFilter on large workbooks when performing statistical operations.

using Aspose.Cells;
using System;

// Demonstrates how to create LoadOptions with a LoadFilter (LoadDataFilterOptions.CellNumeric) to load only numeric and date cells from "Template.xlsx", compute a column sum, and save the result as "Processed.xlsx". This approach reduces memory usage and speeds up statistical calculations.
class Program
{
    static void Main()
    {
        // Create load options and set a filter to load only numeric (including date) cells
        LoadOptions loadOptions = new LoadOptions();
        LoadFilter filter = new LoadFilter(LoadDataFilterOptions.CellNumeric);
        loadOptions.LoadFilter = filter;

        // Load the template workbook with the specified filter
        Workbook workbook = new Workbook("Template.xlsx", loadOptions);

        // Example statistical analysis: sum of numeric values in the first column
        Worksheet sheet = workbook.Worksheets[0];
        double sum = 0;
        int maxRow = sheet.Cells.MaxDataRow;
        for (int row = 0; row <= maxRow; row++)
        {
            if (sheet.Cells[row, 0].Type == CellValueType.IsNumeric)
                sum += sheet.Cells[row, 0].DoubleValue;
        }
        Console.WriteLine($"Sum of numeric values in column A: {sum}");

        // Save the processed workbook
        workbook.Save("Processed.xlsx");
    }
}
