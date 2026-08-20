// Title: C# Custom LoadFilter to Skip Hidden Worksheets in Aspose.Cells
// Description: Shows how to derive a VisibleSheetsLoadFilter from LoadFilter, override StartSheet to load only worksheets with IsVisible = true, attach the filter to LoadOptions, and open a workbook so hidden sheets are omitted. The sample prints loaded sheet names and saves the filtered workbook.
// Keywords: Aspose.Cells LoadFilter C# | skip hidden worksheets Aspose.Cells | load only visible sheets .NET | LoadOptions custom filter | visible worksheets only | Aspose.Cells performance optimization | C# Excel hidden sheet filter
// Common Searches: Aspose.Cells load only visible worksheets | ignore hidden sheets when opening workbook Aspose.Cells | custom LoadFilter example C# | assign LoadFilter to LoadOptions Aspose.Cells | skip hidden worksheets during workbook load .NET
// Developer Intent: Open an Excel workbook while excluding any hidden worksheets by using a custom LoadFilter.
// Use Cases: Reduce memory and processing time when working with large workbooks that contain many hidden tabs. | Create reports that include only user‑visible data, automatically discarding hidden sheets. | Validate which sheets were loaded after applying the filter by iterating workbook.Worksheets.
// AI Prompts: Generate a C# snippet that defines a LoadFilter subclass to skip hidden worksheets and uses it with LoadOptions to open an Excel file in Aspose.Cells. | Explain how to extend VisibleSheetsLoadFilter to also exclude sheets whose names match a specific pattern. | Provide guidance for handling a scenario where all worksheets are hidden and a custom LoadFilter is applied in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterDemo
{
    // Custom filter that loads only visible worksheets
    // Shows how to derive a VisibleSheetsLoadFilter from LoadFilter, override StartSheet to load only worksheets with IsVisible = true, attach the filter to LoadOptions, and open a workbook so hidden sheets are omitted. The sample prints loaded sheet names and saves the filtered workbook.
    public class VisibleSheetsLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load data only if the sheet is visible
            if (sheet.IsVisible)
            {
                // Use default loading behavior for visible sheets
                base.StartSheet(sheet);
            }
            // If the sheet is hidden, do nothing – it will be skipped
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook
            string sourcePath = "InputWorkbook.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new VisibleSheetsLoadFilter();

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // (Optional) Verify which sheets were loaded
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
            }

            // Save the workbook after loading (if needed)
            string outputPath = "FilteredWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}
