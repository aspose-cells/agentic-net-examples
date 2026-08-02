// Title: C# – Load an Excel workbook with Aspose.Cells using LoadOptions (LoadDataOnly = true) to read only cell values
// Description: Demonstrates how to create a LoadOptions object, enable LoadDataOnly (or use LoadFilter CellValue) so that Aspose.Cells reads only raw cell values from an .xlsx file, skipping formatting, formulas, and styles. The example prints the value of cell A1 after the lightweight load.
// Keywords: Aspose.Cells LoadOptions LoadDataOnly | C# read Excel values only | LoadFilter CellValue example | import workbook without formatting | fast Excel data extraction C# | skip styles Aspose.Cells
// Common Searches: Aspose.Cells load only cell values C# | LoadOptions LoadDataOnly true example | How to ignore formatting when opening Excel with Aspose.Cells | C# read Excel data without formulas Aspose | Fast load of large Excel file using Aspose.Cells
// Developer Intent: Open an Excel file with Aspose.Cells while importing just the raw cell contents, omitting all formatting, formulas, and style information.
// Use Cases: Rapid data extraction from massive spreadsheets where visual styles are irrelevant. | Feeding plain values into analytics or ETL pipelines without the overhead of style processing. | Generating lightweight reports or CSV exports that require only the underlying text and numbers.
// AI Prompts: Provide a C# snippet that uses Aspose.Cells LoadOptions with LoadDataOnly = true to read only cell values from an .xlsx file. | Explain the difference between LoadOptions.LoadDataOnly and LoadOptions.LoadFilter when loading Excel workbooks in Aspose.Cells. | Show how to iterate over a range of cells after loading a workbook with LoadFilter set to CellValue.

using System;
using Aspose.Cells;

// Demonstrates how to create a LoadOptions object, enable LoadDataOnly (or use LoadFilter CellValue) so that Aspose.Cells reads only raw cell values from an .xlsx file, skipping formatting, formulas, and styles. The example prints the value of cell A1 after the lightweight load.
class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only cell values (no formatting, formulas, etc.)
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellValue);

        // Load the workbook using the configured options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Example usage: output the value of cell A1 from the first worksheet
        Console.WriteLine(workbook.Worksheets[0].Cells["A1"].StringValue);
    }
}
