using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Prepare load options with a custom filter that selects sheets named like "Q1", "Q2", etc.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new QSheetLoadFilter();

        // Load the workbook using the filter
        Workbook workbook = new Workbook("Template.xlsx", loadOptions);

        // Save the workbook containing only the selected worksheets
        workbook.Save("FilteredWorkbook.xlsx");
    }

    // Custom LoadFilter implementation
    private class QSheetLoadFilter : LoadFilter
    {
        // Pre‑compiled regular expression for performance
        private static readonly Regex qSheetRegex = new Regex(@"^Q\d+$", RegexOptions.Compiled);

        public override void StartSheet(Worksheet sheet)
        {
            // If the worksheet name matches "Q[0-9]+", load all its data
            if (qSheetRegex.IsMatch(sheet.Name))
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            // Otherwise load only the structure (no cell data), effectively excluding it from the result
            {
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }
}

// Author: Aspose.Cells .NET example – loads only worksheets matching "Q[0-9]+" using LoadFilter.