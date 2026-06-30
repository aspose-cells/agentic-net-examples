using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – loads a workbook while skipping hidden worksheets
    public class CustomLoadFilter : LoadFilter
    {
        // Called before each worksheet is loaded
        public override void StartSheet(Worksheet sheet)
        {
            // Load only if the worksheet is visible
            if (sheet.IsVisible)
            {
                base.StartSheet(sheet);
            }
            // If the sheet is hidden, do not call base.StartSheet – it will be skipped
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with your actual file)
            string sourcePath = "InputWorkbook.xlsx";

            // Configure LoadOptions with the custom filter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new CustomLoadFilter()
            };

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Output the names and visibility of the loaded worksheets
            Console.WriteLine("Loaded worksheets (hidden sheets excluded):");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
                // Example: read a cell value from each loaded sheet
                Console.WriteLine($"  A1 Value: {ws.Cells["A1"].StringValue}");
            }

            // Optionally save the filtered workbook to a new file
            string outputPath = "FilteredWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}