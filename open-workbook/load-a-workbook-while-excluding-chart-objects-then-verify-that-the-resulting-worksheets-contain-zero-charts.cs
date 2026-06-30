using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Author: Aspose.Cells .NET example – load workbook without charts and verify

class Program
{
    static void Main()
    {
        // Create load options; ignore useless shapes (does not affect charts but follows the rule)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.IgnoreUselessShapes = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Remove all chart objects from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Charts.Clear(); // ChartCollection.Clear removes all charts in the worksheet
        }

        // Verify that each worksheet now contains zero charts
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' chart count: {sheet.Charts.Count}");
        }

        // (Optional) Save the workbook to confirm the result
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}