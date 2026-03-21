using System;
using Aspose.Cells;

namespace AsposeCellsLoadWorksheetsDemo
{
    // Custom load filter to demonstrate selective loading of worksheets
    class CustomLoadFilter : LoadFilter
    {
        // Load all data for sheets named "Data", otherwise load only structure
        public override void StartSheet(Worksheet sheet)
        {
            if (sheet.Name == "Data")
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            else
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source Excel file containing multiple worksheets
            string sourceFile = "SourceWorkbook.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter();

            // Load the workbook using the constructor that accepts file path and LoadOptions
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Access the worksheets collection
            WorksheetCollection sheets = workbook.Worksheets;

            // Example: copy an existing worksheet named "Template" to a new sheet called "WorkingCopy"
            // Ensure the source sheet exists before copying
            if (sheets["Template"] != null)
            {
                int copiedIndex = sheets.AddCopy("Template");
                Worksheet copiedSheet = sheets[copiedIndex];
                copiedSheet.Name = "WorkingCopy";
            }

            // Perform any additional data manipulation here...
            // For demonstration, write a timestamp into the first cell of the first worksheet
            sheets[0].Cells["A1"].PutValue($"Loaded on {DateTime.Now}");

            // Save the modified workbook to a new file
            workbook.Save("ProcessedWorkbook.xlsx");
        }
    }
}