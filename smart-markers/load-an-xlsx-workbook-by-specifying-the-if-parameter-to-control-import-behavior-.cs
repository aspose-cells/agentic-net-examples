using System;
using Aspose.Cells;

namespace AsposeCellsLoadWithFilterDemo
{
    // Custom filter to control what data is loaded for each worksheet
    public class CustomLoadFilter : LoadFilter
    {
        // This method is called before a worksheet is loaded.
        // We can set the LoadDataFilterOptions here to specify what to import.
        public override void StartSheet(Worksheet sheet)
        {
            // Load all data for "Sheet1", otherwise load only the structure (no cell data)
            if (sheet.Name.Equals("Sheet1", StringComparison.OrdinalIgnoreCase))
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter();

            // Load the workbook with the specified options (IF parameter = LoadFilter)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook after loading with the applied filter
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}