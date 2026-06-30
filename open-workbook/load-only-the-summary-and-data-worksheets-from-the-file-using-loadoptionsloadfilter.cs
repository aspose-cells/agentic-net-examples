using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterDemo
{
    // Author: Aspose.Cells .NET example
    // Custom LoadFilter to load only "Summary" and "Data" worksheets with full data.
    class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data for the required sheets, otherwise load only the structure.
            if (sheet.Name == "Summary" || sheet.Name == "Data")
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                // Load only the sheet structure (no cell data) for other sheets.
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook.
            string sourcePath = "TemplateWorkbook.xlsx";

            // Configure LoadOptions with the custom filter.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new CustomLoadFilter()
            };

            // Load the workbook using the specified LoadOptions.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // (Optional) Save the filtered workbook to verify the result.
            workbook.Save("FilteredWorkbook.xlsx");
        }
    }
}