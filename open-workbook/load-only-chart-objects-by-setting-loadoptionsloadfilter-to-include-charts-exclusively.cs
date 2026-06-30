using System;
using Aspose.Cells;

public class LoadChartsOnlyDemo
{
    public static void Main()
    {
        // Configure LoadOptions to use a custom filter that loads only charts
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new ChartOnlyLoadFilter();

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook; only chart objects are loaded
        workbook.Save("output.xlsx");
    }

    // Custom LoadFilter that sets the filter to load only chart objects per worksheet
    private class ChartOnlyLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only charts for the current sheet
            LoadDataFilterOptions = LoadDataFilterOptions.Chart;
        }
    }
}

// Author: Aspose.Cells .NET example – loads only chart objects.