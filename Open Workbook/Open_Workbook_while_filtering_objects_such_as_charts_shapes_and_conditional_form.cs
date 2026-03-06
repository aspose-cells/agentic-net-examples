using System;
using Aspose.Cells;

class LoadWorkbookWithFilteredObjects
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Create LoadOptions and set a LoadFilter that excludes charts, shapes, and conditional formatting
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(
            // Start with loading everything, then remove the unwanted object types
            LoadDataFilterOptions.All
            & ~LoadDataFilterOptions.Chart          // Exclude charts
            & ~LoadDataFilterOptions.Shape          // Exclude shapes
            & ~LoadDataFilterOptions.ConditionalFormatting // Exclude conditional formatting
        );

        // Load the workbook using the specified load options
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Save the filtered workbook to a new file
        workbook.Save("output.xlsx");
    }
}