// Title: C# – Load only cell data (skip charts) with Aspose.Cells LoadFilter to reduce memory usage
// Description: Demonstrates how to create a LoadOptions object, set its LoadFilter to LoadDataFilterOptions.CellData, and open an Excel workbook so that only cell values, formulas, and formatting are loaded while chart objects are omitted, resulting in lower memory consumption. The workbook is then saved to confirm the data load.
// Keywords: Aspose.Cells | LoadFilter | LoadOptions | CellData | skip charts | memory optimization | .NET | C# | Excel workbook loading | load only cells | chart‑free loading
// Common Searches: How to open an Excel file without loading charts using Aspose.Cells C# | Aspose.Cells LoadFilter example to load only cell data | Reduce memory usage when loading workbooks with charts Aspose.Cells | Load workbook with cell values only Aspose.Cells .NET | Skip chart objects on workbook load Aspose.Cells
// Developer Intent: Open an Excel workbook while excluding chart objects to minimize memory consumption.
// Use Cases: Process large reporting workbooks that contain many charts but require only cell values for calculations. | Extract raw data from chart‑heavy templates without rendering graphics, keeping the operation lightweight. | Run batch data validation on thousands of workbooks in a cloud service while keeping the memory footprint low. | Generate CSV or JSON exports from financial models that embed numerous charts, without loading the chart data.
// AI Prompts: Show code to also exclude images and shapes using LoadFilter. | Provide an example of loading only formulas and formatting while omitting cell values. | Explain how to verify which object types were skipped after loading a workbook with LoadFilter. | Give a performance comparison of loading a workbook with and without charts using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Demonstrates how to create a LoadOptions object, set its LoadFilter to LoadDataFilterOptions.CellData, and open an Excel workbook so that only cell values, formulas, and formatting are loaded while chart objects are omitted, resulting in lower memory consumption. The workbook is then saved to confirm the data load.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that may contain charts
            string sourcePath = "TemplateWithCharts.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Configure LoadFilter to load only cell data (values, formulas, formatting) and skip charts
            // LoadDataFilterOptions.CellData includes cell values, formulas and formatting but excludes charts
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellData);

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point, charts are not loaded into the workbook, reducing memory usage

            // Save the workbook to verify that data cells are loaded correctly
            string outputPath = "LoadedDataOnly.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook loaded with only cell data and saved to '{outputPath}'.");
        }
    }
}
