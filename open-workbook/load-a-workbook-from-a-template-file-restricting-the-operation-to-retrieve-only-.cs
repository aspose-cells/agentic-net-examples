using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create LoadOptions and configure it to load only shape objects
        LoadOptions loadOptions = new LoadOptions();
        // LoadFilter with LoadDataFilterOptions.Shape ensures only shapes are loaded
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Shape);
        // Optional: ignore overlapping useless shapes for better performance
        loadOptions.IgnoreUselessShapes = true;

        // Load the workbook from the template file using the specified options
        Workbook workbook = new Workbook("Template.xlsx", loadOptions);

        // Example: iterate through worksheets and output the number of loaded shapes
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            Console.WriteLine($"Worksheet '{sheet.Name}' contains {sheet.Shapes.Count} shape(s).");
        }

        // Save the workbook if further processing or saving is required
        workbook.Save("Result.xlsx");
    }
}