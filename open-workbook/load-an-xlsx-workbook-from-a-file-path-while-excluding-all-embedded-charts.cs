using System;
using Aspose.Cells;

class LoadWorkbookWithoutCharts
{
    static void Main()
    {
        // Path to the source XLSX file
        string filePath = "input.xlsx";

        // Create LoadOptions and configure a LoadFilter that loads everything except charts
        LoadOptions loadOptions = new LoadOptions();
        // Load all data but exclude charts (Chart flag = 256)
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // At this point the workbook is loaded without any embedded charts
        Console.WriteLine("Workbook loaded. Worksheet count: " + workbook.Worksheets.Count);
        // Example: display number of shapes (charts are shapes) in the first worksheet
        Console.WriteLine("Shapes in first worksheet: " + workbook.Worksheets[0].Shapes.Count);
    }
}