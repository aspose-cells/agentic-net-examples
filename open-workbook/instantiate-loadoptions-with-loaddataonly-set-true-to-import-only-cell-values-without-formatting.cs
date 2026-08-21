// Title: C# – Load an Excel workbook with Aspose.Cells using LoadOptions to read only cell values (no formatting)
// Description: Demonstrates how to create a LoadOptions object, set its LoadFilter to LoadDataFilterOptions.CellValue, and open an Excel file so that only raw cell values are loaded while all formatting, styles, and graphics are ignored. The example prints the value of cell A1 from the first worksheet, showing a lightweight way to read data.
// Keywords: Aspose.Cells LoadOptions C# | LoadDataOnly true Aspose.Cells | LoadFilter CellValue example | read Excel values only | skip formatting Aspose.Cells | memory‑efficient Excel import C# | load workbook without styles
// Common Searches: Aspose.Cells load workbook without formatting C# | LoadOptions LoadDataOnly true example | How to read only cell values with Aspose.Cells | C# load Excel file ignoring styles | LoadFilter CellValue usage Aspose.Cells
// Developer Intent: Open an Excel file with Aspose.Cells while retrieving only the raw cell values and omitting all formatting, styles, and graphics.
// Use Cases: Fast extraction of data from large spreadsheets for analytics or ETL pipelines. | Migrating cell contents to a database without importing visual formatting. | Running batch validation or calculations on cell values with minimal memory overhead.
// AI Prompts: Show how to modify the code to also load formulas while still skipping formatting. | Provide a version that reads only cell values from a CSV file using Aspose.Cells LoadOptions. | Explain how to combine LoadOptions with LoadFilter to exclude comments and hyperlinks.

using System;
using Aspose.Cells;

// Demonstrates how to create a LoadOptions object, set its LoadFilter to LoadDataFilterOptions.CellValue, and open an Excel file so that only raw cell values are loaded while all formatting, styles, and graphics are ignored. The example prints the value of cell A1 from the first worksheet, showing a lightweight way to read data.
class LoadDataOnlyDemo
{
    static void Main()
    {
        // Create LoadOptions and configure it to load only cell values (no formatting)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

        // Load the workbook using the configured options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet and display a cell value to verify loading
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);
    }
}
