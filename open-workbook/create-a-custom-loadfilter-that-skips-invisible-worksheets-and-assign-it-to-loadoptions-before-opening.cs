using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterDemo
{
    // Custom LoadFilter that loads only visible worksheets
    public class SkipInvisibleSheetsFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load data only when the worksheet is visible
            if (sheet.IsVisible)
            {
                // Proceed with default loading behavior
                base.StartSheet(sheet);
            }
            // If the sheet is hidden, do nothing – it will be skipped
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the workbook to be loaded
            string inputFile = "InputWorkbook.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new SkipInvisibleSheetsFilter();

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Display information about the loaded worksheets
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
            }

            // Save the workbook after filtering (optional)
            string outputFile = "FilteredWorkbook.xlsx";
            workbook.Save(outputFile);
        }
    }
}