using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – custom LoadFilter to skip invisible worksheets
public class CustomLoadFilter : LoadFilter
{
    // This method is called for each worksheet during loading.
    public override void StartSheet(Worksheet sheet)
    {
        // Load full data only for visible sheets.
        if (sheet.IsVisible)
        {
            LoadDataFilterOptions = LoadDataFilterOptions.All;
        }
        else
        {
            // For invisible sheets load only the structure (no cell data).
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}

class Program
{
    static void Main()
    {
        // Path to the source workbook.
        string sourcePath = "input.xlsx";

        // Create LoadOptions and assign the custom filter.
        LoadOptions options = new LoadOptions();
        options.LoadFilter = new CustomLoadFilter();

        // Load the workbook using the configured LoadOptions.
        Workbook workbook = new Workbook(sourcePath, options);

        // Save the result to verify that invisible sheets were skipped.
        workbook.Save("output.xlsx");
    }
}