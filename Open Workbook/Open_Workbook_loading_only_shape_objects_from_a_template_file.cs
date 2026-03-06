using System;
using Aspose.Cells;

class LoadOnlyShapesDemo
{
    static void Main()
    {
        // Path to the template workbook
        string templatePath = "template.xlsx";

        // Create LoadOptions and configure it to load only shape objects
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Shape);

        // Load the workbook using the constructor that accepts a file path and LoadOptions
        Workbook workbook = new Workbook(templatePath, loadOptions);

        // Retrieve the first worksheet and output the count of loaded shapes
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Number of shapes loaded: " + sheet.Shapes.Count);
    }
}