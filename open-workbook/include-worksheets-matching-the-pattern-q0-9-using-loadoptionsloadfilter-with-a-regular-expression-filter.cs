// Title: C# – Load only worksheets matching "Q[0-9]+" with Aspose.Cells LoadOptions and a Regex LoadFilter
// Description: Demonstrates how to create a Regex‑based LoadFilter that loads full data for worksheets whose names match the pattern ^Q\d+$ and loads only the structure for all other sheets. The filter is assigned to LoadOptions, used to open an Excel file, and the filtered workbook is saved, reducing memory consumption and processing time.
// Keywords: Aspose.Cells | C# | .NET | LoadOptions | LoadFilter | RegexLoadFilter | selective worksheet loading | worksheet name pattern | memory optimization | Excel sheet filter | Q1 Q2 Q3 worksheets
// Common Searches: Aspose.Cells load worksheets by name pattern | C# regex LoadFilter for Excel files | Load only sheets starting with Q in Aspose.Cells | How to skip sheet data with LoadDataFilterOptions | Selective sheet loading using LoadOptions .NET
// Developer Intent: Load a workbook while including full data only for sheets whose names match Q[0-9]+ and keep the rest as structure‑only placeholders.
// Use Cases: Extract quarterly sheets (Q1, Q2, …) from a massive workbook without loading unrelated data. | Create a lightweight copy that contains only question‑type worksheets for reporting or distribution. | Improve performance and lower memory usage by loading non‑matching sheets in structure‑only mode.
// AI Prompts: Write a C# example that uses Aspose.Cells LoadOptions with a custom Regex LoadFilter to load only worksheets named with the pattern Q[0-9]+. | Explain the effect of LoadDataFilterOptions.Structure on sheets that do not match the regex in a custom LoadFilter. | Show how to extend the RegexLoadFilter to accept multiple patterns, such as Q[0-9]+ and "Summary".

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Demonstrates how to create a Regex‑based LoadFilter that loads full data for worksheets whose names match the pattern ^Q\d+$ and loads only the structure for all other sheets. The filter is assigned to LoadOptions, used to open an Excel file, and the filtered workbook is saved, reducing memory consumption and processing time.
class Program
{
    static void Main()
    {
        // Create a custom load filter that includes only worksheets whose names match "Q[0-9]+"
        LoadFilter filter = new RegexLoadFilter(@"^Q\d+$");

        // Set the filter in LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = filter;

        // Load the workbook using the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook after loading the filtered sheets
        workbook.Save("output.xlsx");
    }

    // Custom LoadFilter implementation using a regular expression
    class RegexLoadFilter : LoadFilter
    {
        private readonly Regex _namePattern;

        public RegexLoadFilter(string pattern)
        {
            _namePattern = new Regex(pattern, RegexOptions.Compiled);
        }

        public override void StartSheet(Worksheet sheet)
        {
            // If the worksheet name matches the pattern, load all its data;
            // otherwise, load only the structure (effectively skipping the sheet's content)
            if (_namePattern.IsMatch(sheet.Name))
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            else
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}
