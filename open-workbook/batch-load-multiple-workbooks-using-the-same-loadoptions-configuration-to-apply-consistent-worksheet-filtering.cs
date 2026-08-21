// Title: Batch load Excel workbooks with a shared LoadOptions and custom LoadFilter using Aspose.Cells for .NET
// Description: Demonstrates how to create a single LoadOptions object that contains a CustomLoadFilter, then reuse it to open multiple workbooks in a loop. The filter loads full cell data only for worksheets whose names start with "Data" and loads just the structure for all other sheets, reducing memory usage and speeding up batch processing.
// Keywords: Aspose.Cells batch loading | shared LoadOptions | custom LoadFilter .NET | selective worksheet loading | load worksheet structure only | C# Excel performance | multiple workbook processing
// Common Searches: reuse LoadOptions for several workbooks Aspose.Cells | load only specific sheets data in batch with Aspose.Cells | apply custom LoadFilter to multiple Excel files .NET | how to improve performance when opening many workbooks Aspose.Cells
// Developer Intent: Open many Excel files with one LoadOptions instance that contains a custom LoadFilter, controlling per‑sheet data loading to optimize speed and memory consumption.
// Use Cases: Extract summary information from dozens of report files while skipping heavy data in non‑report sheets. | Build an ETL pipeline that reads a batch of workbooks, modifies only the "Data*" sheets, and writes the files back. | Generate a quick inventory of worksheet counts across a folder of workbooks without loading full cell contents.
// AI Prompts: Show how to extend CustomLoadFilter to also ignore charts and images on non‑Data worksheets. | Provide a Parallel.ForEach example that loads workbooks concurrently while sharing the same LoadOptions. | Create logging code that records which sheets were loaded with full data versus structure only during batch processing.

using System;
using Aspose.Cells;

// Demonstrates how to create a single LoadOptions object that contains a CustomLoadFilter, then reuse it to open multiple workbooks in a loop. The filter loads full cell data only for worksheets whose names start with "Data" and loads just the structure for all other sheets, reducing memory usage and speeding up batch processing.
class CustomLoadFilter : LoadFilter
{
    // Adjust loading options per worksheet
    public override void StartSheet(Worksheet sheet)
    {
        // Load full data for sheets whose name starts with "Data"
        // Otherwise load only the worksheet structure
        if (sheet.Name.StartsWith("Data", StringComparison.OrdinalIgnoreCase))
            LoadDataFilterOptions = LoadDataFilterOptions.All;
        else
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
    }
}

class Program
{
    static void Main()
    {
        // Create a single LoadOptions instance and assign the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter();

        // Paths of workbooks to be loaded in batch
        string[] workbookFiles = { "Book1.xlsx", "Book2.xlsx", "Book3.xlsx" };

        foreach (string filePath in workbookFiles)
        {
            // Load each workbook using the shared LoadOptions configuration
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example: display the number of worksheets loaded
            Console.WriteLine($"'{filePath}' loaded with {workbook.Worksheets.Count} worksheets.");

            // Save the workbook to verify successful loading (optional)
            string outputPath = System.IO.Path.GetFileNameWithoutExtension(filePath) + "_processed.xlsx";
            workbook.Save(outputPath);
        }
    }
}
