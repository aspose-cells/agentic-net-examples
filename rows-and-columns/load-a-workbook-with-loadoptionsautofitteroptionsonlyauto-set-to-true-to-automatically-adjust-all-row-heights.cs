using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create LoadOptions and configure AutoFitterOptions to only auto‑fit rows whose height is not custom set
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = new AutoFitterOptions
        {
            OnlyAuto = true
        };

        // Load the workbook with the specified options; rows will be auto‑fitted during load
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Example: display the height of the first row after auto‑fit
        double rowHeight = workbook.Worksheets[0].Cells.GetRowHeight(0);
        Console.WriteLine($"Row 0 height after auto‑fit: {rowHeight}");

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}

// Author: Aspose.Cells .NET example