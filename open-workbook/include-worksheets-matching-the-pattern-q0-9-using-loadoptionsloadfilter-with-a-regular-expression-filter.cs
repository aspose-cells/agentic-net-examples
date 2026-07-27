// Title: C# – Load Worksheets Matching "Q[0-9]+" Using Aspose.Cells LoadFilter (Regex)
// Description: Demonstrates a custom RegexLoadFilter that inherits from LoadFilter, applies a pre‑compiled ^Q\d+$ pattern, and sets LoadDataFilterOptions.All for matching sheets while skipping others. The filter is assigned to LoadOptions, the workbook is opened, loaded sheet names are displayed, and the filtered workbook can be saved.
// Keywords: Aspose.Cells | LoadFilter | C# | .NET | regular expression | worksheet filter | LoadOptions | regex sheet selection | performance optimization | partial workbook load
// Common Searches: Aspose.Cells load only specific sheets | C# LoadFilter regex example | filter worksheets by name when opening workbook Aspose.Cells | custom LoadFilter with Aspose.Cells | exclude sheets not matching pattern Aspose.Cells
// Developer Intent: Open a workbook while loading exclusively the worksheets whose names satisfy the pattern Q[0-9]+.
// Use Cases: Extract quarterly tabs (Q1, Q2, …) from a large financial workbook for targeted analysis. | Create a lightweight copy that contains only question sheets for downstream reporting. | Reduce memory consumption and load time by skipping unrelated worksheets in massive Excel files.
// AI Prompts: Generate a C# LoadFilter that includes worksheets based on any user‑provided regex pattern using Aspose.Cells. | Show how to invert the RegexLoadFilter logic to exclude matching sheets instead of including them. | Illustrate combining LoadFilter with LoadDataFilterOptions to load only the first 100 rows of matching worksheets.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Custom LoadFilter that includes only worksheets whose names match the pattern "Q[0-9]+"
    // Demonstrates a custom RegexLoadFilter that inherits from LoadFilter, applies a pre‑compiled ^Q\d+$ pattern, and sets LoadDataFilterOptions.All for matching sheets while skipping others. The filter is assigned to LoadOptions, the workbook is opened, loaded sheet names are displayed, and the filtered workbook can be saved.
    class RegexLoadFilter : LoadFilter
    {
        // Pre‑compiled regular expression for performance
        private static readonly Regex SheetNamePattern = new Regex(@"^Q\d+$", RegexOptions.Compiled);

        public override void StartSheet(Worksheet sheet)
        {
            // If the sheet name matches the pattern, load all its data;
            // otherwise load nothing (the sheet will be effectively excluded)
            if (SheetNamePattern.IsMatch(sheet.Name))
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            else
                LoadDataFilterOptions = LoadDataFilterOptions.None;
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare load options and assign the custom filter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new RegexLoadFilter()
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook("Template.xlsx", loadOptions);

            // Display the names of the worksheets that were actually loaded
            Console.WriteLine("Worksheets loaded after applying the regex filter:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine("- " + ws.Name);
            }

            // Save the filtered workbook (optional)
            workbook.Save("FilteredOutput.xlsx");
        }
    }
}
