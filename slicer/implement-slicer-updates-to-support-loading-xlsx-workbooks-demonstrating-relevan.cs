using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class SlicerLoadDemo
{
    static void Main()
    {
        // Create a custom LoadFilter to control what data is loaded for each sheet
        var loadFilter = new CustomLoadFilter();

        // Assign the filter to LoadOptions
        var loadOptions = new LoadOptions();
        loadOptions.LoadFilter = loadFilter;

        // Load the workbook from an existing XLSX file using the specified LoadOptions
        var workbook = new Workbook("TemplateWithPivot.xlsx", loadOptions);

        // Assume the first worksheet holds the source data and the second contains the pivot table and slicer
        var dataSheet = workbook.Worksheets[0];
        var pivotSheet = workbook.Worksheets[1];

        // Modify source data to demonstrate that the slicer will reflect the change after refresh
        dataSheet.Cells["B2"].PutValue(999); // Example: change a sales figure

        // Refresh all slicers on the pivot sheet; this also refreshes the underlying pivot tables
        var slicers = pivotSheet.Slicers;
        for (int i = 0; i < slicers.Count; i++)
        {
            slicers[i].Refresh();
        }

        // Save the updated workbook
        workbook.Save("FilteredSlicerDemo.xlsx");
    }

    // Custom LoadFilter implementation
    class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data (including values, formulas, charts, slicers) for sheets that contain pivot tables
            // Here we identify such sheets by name containing "Pivot"
            if (sheet.Name.Contains("Pivot"))
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                // For other sheets, load only the structure to speed up loading
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }
}