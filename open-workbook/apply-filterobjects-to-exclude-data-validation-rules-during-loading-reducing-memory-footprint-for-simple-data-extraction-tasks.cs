// Title: C# – Load an Excel workbook with Aspose.Cells while skipping data validation to save memory
// Description: Shows how to configure Aspose.Cells LoadOptions and LoadFilter in C# to exclude DataValidation objects when opening a workbook. The filter clears the DataValidation flag, preserving cells, formulas, and formatting, enabling low‑memory extraction such as reading A1 values and optionally saving the file.
// Keywords: Aspose.Cells LoadFilter | LoadOptions DataValidation | skip data validation | reduce memory usage | C# Excel loading | exclude data validation Aspose | memory‑efficient workbook load | read cell values without validation | Aspose.Cells performance | filter workbook load
// Common Searches: Aspose.Cells load workbook without data validation | C# LoadFilter exclude DataValidation | how to reduce memory when loading Excel with Aspose | skip data validation objects Aspose.Cells | LoadOptions to ignore validation rules .NET
// Developer Intent: Load a workbook while omitting data validation objects to lower memory consumption.
// Use Cases: Processing large Excel files when only cell values are required. | Performing quick data extraction (e.g., reading A1 from each sheet) without the overhead of validation rules. | Running batch jobs on limited‑memory servers or containers. | Saving a workbook after read‑only operations to confirm that filtered loading does not affect functionality.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadFilter to open an Excel file without loading data validation objects. | Explain how to combine LoadFilter with other LoadDataFilterOptions for custom workbook loading. | Provide a step‑by‑step guide to read cell values from each worksheet after excluding data validation with LoadOptions.

using System;
using Aspose.Cells;

// Shows how to configure Aspose.Cells LoadOptions and LoadFilter in C# to exclude DataValidation objects when opening a workbook. The filter clears the DataValidation flag, preserving cells, formulas, and formatting, enabling low‑memory extraction such as reading A1 values and optionally saving the file.
class Program
{
    static void Main()
    {
        // Paths for the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Create a LoadFilter with default options (LoadDataFilterOptions.All)
        LoadFilter filter = new LoadFilter();

        // Remove the DataValidation flag from the filter options
        // This keeps all other data (cells, formulas, formatting, etc.) but skips data validations
        filter.LoadDataFilterOptions = filter.LoadDataFilterOptions & ~LoadDataFilterOptions.DataValidation;

        // Assign the customized filter to the load options
        loadOptions.LoadFilter = filter;

        // Load the workbook using the filter to reduce memory usage
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Simple extraction: display the value of cell A1 from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' - A1: {sheet.Cells["A1"].StringValue}");
        }

        // Save the workbook (optional, demonstrates that the workbook remains functional)
        workbook.Save(outputPath);
    }
}
