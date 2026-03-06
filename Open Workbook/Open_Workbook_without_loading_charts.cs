using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create LoadOptions and assign a custom LoadFilter that excludes charts
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new SkipChartLoadFilter();

        // Load the workbook using the options – charts will not be loaded
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Verify that no charts were loaded (should be 0)
        Console.WriteLine("Number of charts loaded: " + workbook.Worksheets[0].Charts.Count);

        // Save the workbook (charts are absent)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }

    // Custom LoadFilter implementation that loads everything except charts
    private class SkipChartLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load all data except charts
            this.LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        }
    }
}