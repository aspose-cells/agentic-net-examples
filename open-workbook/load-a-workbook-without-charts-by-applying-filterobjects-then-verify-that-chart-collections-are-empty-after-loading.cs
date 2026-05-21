using System;
using Aspose.Cells;

class LoadWorkbookWithoutCharts
{
    static void Main()
    {
        // Path to the source workbook that contains charts
        string sourcePath = "ChartWorkbook.xlsx";

        // Create a LoadFilter that loads only cell data (excluding charts)
        LoadFilter filter = new LoadFilter();
        filter.LoadDataFilterOptions = LoadDataFilterOptions.CellData; // Load cells, formulas, formatting but not charts

        // Assign the filter to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = filter;

        // Load the workbook with the specified load options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Verify that each worksheet's chart collection is empty
        bool allChartsRemoved = true;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Charts.Count > 0)
            {
                allChartsRemoved = false;
                Console.WriteLine($"Worksheet '{sheet.Name}' still contains {sheet.Charts.Count} chart(s).");
            }
        }

        if (allChartsRemoved)
        {
            Console.WriteLine("All worksheets loaded without charts.");
        }

        // Optionally save the workbook to confirm the result
        workbook.Save("ChartWorkbook_NoCharts.xlsx");
    }
}