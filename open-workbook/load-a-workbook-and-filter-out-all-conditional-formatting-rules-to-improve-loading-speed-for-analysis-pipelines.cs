using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Custom filter that disables loading of conditional formatting rules
    public class NoConditionalFormattingFilter : LoadFilter
    {
        // Override to set filter options for each worksheet during loading
        public override void StartSheet(Worksheet sheet)
        {
            // Load everything except conditional formatting
            // All = 2147483647, ConditionalFormatting = 2048
            // Use bitwise AND with the complement of ConditionalFormatting flag
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.ConditionalFormatting;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new NoConditionalFormattingFilter();

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // (Optional) Save the workbook to verify that conditional formatting is not loaded
            workbook.Save("output_without_conditional_formatting.xlsx");
        }
    }
}