using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom filter that loads full data only for selected worksheets.
    public class SelectedSheetsLoadFilter : LoadFilter
    {
        // Names of worksheets that should be fully loaded.
        private readonly string[] _sheetsToLoad;

        public SelectedSheetsLoadFilter(string[] sheetsToLoad)
        {
            _sheetsToLoad = sheetsToLoad ?? Array.Empty<string>();
        }

        // This method is called for each worksheet before it is loaded.
        public override void StartSheet(Worksheet sheet)
        {
            // If the sheet name is in the list, load everything; otherwise load only the structure.
            if (Array.IndexOf(_sheetsToLoad, sheet.Name) >= 0)
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    public class LoadWorkbookWithSelectedSheets
    {
        public static void Run()
        {
            // Path to the source Excel file.
            string sourceFile = "input.xlsx";

            // Define which worksheets should be fully initialized.
            string[] sheetsToInitialize = { "Sheet2", "Sheet4" };

            // Create LoadOptions and assign the custom LoadFilter.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new SelectedSheetsLoadFilter(sheetsToInitialize);

            // Load the workbook using the constructor that accepts a file path and LoadOptions.
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // At this point only Sheet2 and Sheet4 contain full cell data;
            // other sheets contain only their structure (no cell values, formulas, etc.).

            // Example: iterate through loaded worksheets and display basic info.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {ws.Name}");
                Console.WriteLine($"  Cells count (may be zero if only structure loaded): {ws.Cells.Count}");
            }

            // Save the workbook to a new file to verify the result.
            workbook.Save("output_selected_sheets.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookWithSelectedSheets.Run();
        }
    }
}