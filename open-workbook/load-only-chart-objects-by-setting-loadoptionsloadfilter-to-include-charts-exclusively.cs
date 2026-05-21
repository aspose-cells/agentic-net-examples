using System;
using Aspose.Cells;

class LoadChartsOnly
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only chart objects
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Chart);

        // Load the workbook with the specified filter
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Display the number of charts loaded in each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' contains {sheet.Charts.Count} chart(s).");
        }

        // Save the workbook (it will contain only the loaded charts)
        workbook.Save("output_charts_only.xlsx");
    }
}