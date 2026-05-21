using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create LoadOptions and assign a custom LoadFilter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter();

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the loaded workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }

    // Custom LoadFilter to load only sheets 0 and 2 and exclude charts
    private class CustomLoadFilter : LoadFilter
    {
        // Load only the worksheets with indexes 0 and 2
        public override int[] SheetsInLoadingOrder => new int[] { 0, 2 };

        // Before each sheet is loaded, set filter options to load everything except charts
        public override void StartSheet(Worksheet sheet)
        {
            // Load all data but remove the Chart flag
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        }
    }
}