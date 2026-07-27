// Title: Load an Excel workbook with Aspose.Cells while skipping data validation objects to cut memory usage (C#)
// Description: Demonstrates how to create a LoadFilter that removes the DataValidation flag from LoadDataFilterOptions, assign it to LoadOptions, and open a workbook without loading any validation rules. The example verifies that no validations are present and extracts raw cell values for lightweight data extraction.
// Keywords: Aspose.Cells LoadFilter | exclude data validation | reduce memory footprint Excel | LoadDataFilterOptions DataValidation | C# read Excel without validations | fast Excel data extraction | memory‑efficient workbook loading | skip validation objects Aspose
// Common Searches: How to load an Excel file with Aspose.Cells and ignore data validations | Aspose.Cells LoadFilter example for memory optimization | Skip data validation objects when opening a workbook in .NET | LoadDataFilterOptions to exclude validations in Aspose.Cells | Improve performance of Excel reading by omitting validation rules
// Developer Intent: Open a workbook with Aspose.Cells while omitting data validation objects to lower memory consumption for simple extraction tasks.
// Use Cases: Reading large spreadsheets for reporting where only cell values matter. | ETL pipelines that migrate raw data without validation metadata. | Automated tests that verify content without the overhead of validation collections.
// AI Prompts: Show how to extend the LoadFilter to also skip comments, shapes, and charts. | Provide a LoadFilter configuration that loads only formulas and values, excluding validations and graphics. | Explain how to benchmark memory savings when using LoadFilter to ignore data validations in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    // Demonstrates how to create a LoadFilter that removes the DataValidation flag from LoadDataFilterOptions, assign it to LoadOptions, and open a workbook without loading any validation rules. The example verifies that no validations are present and extracts raw cell values for lightweight data extraction.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "sample.xlsx";

            // Create a LoadFilter that loads everything except data validations
            // LoadDataFilterOptions.All includes all data; we remove the DataValidation flag using bitwise NOT
            LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.DataValidation;
            LoadFilter loadFilter = new LoadFilter(filterOptions);

            // Assign the filter to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = loadFilter;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Demonstrate that data validations are not loaded
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' has {sheet.Validations.Count} data validations loaded.");
            }

            // Example: read and display some cell values (simple data extraction)
            Worksheet firstSheet = workbook.Worksheets[0];
            int maxRow = firstSheet.Cells.MaxDataRow;
            int maxCol = firstSheet.Cells.MaxDataColumn;

            Console.WriteLine("\nExtracted cell values:");
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Console.Write($"{firstSheet.Cells[row, col].StringValue}\t");
                }
                Console.WriteLine();
            }

            // No need to save the workbook as we only performed extraction
        }
    }
}
