// Title: Exclude Hidden Worksheets When Loading a Workbook with Aspose.Cells (C#)
// Description: Demonstrates how to configure Aspose.Cells LoadOptions with a custom LoadFilter that loads only visible worksheets. The example creates a VisibleSheetLoadFilter, applies it to LoadOptions, opens the workbook, lists the loaded sheets, and saves the result, reducing memory usage and processing time.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter | visible worksheets | skip hidden sheets | C# Excel loading | exclude hidden tabs | .NET Excel API | performance optimization
// Common Searches: Aspose.Cells load workbook without hidden sheets | How to ignore hidden worksheets in C# Aspose.Cells | Load only visible worksheets using LoadFilter | Skip hidden tabs when opening Excel with Aspose.Cells | Custom LoadFilter example Aspose.Cells .NET
// Developer Intent: Load an Excel file while automatically omitting hidden worksheets.
// Use Cases: Process only visible sheets for reporting or analytics to improve speed. | Create a lightweight copy of a workbook that contains just the visible tabs. | Export or validate data from visible worksheets while ignoring hidden ones.
// AI Prompts: Write a C# snippet that uses Aspose.Cells LoadOptions with a custom LoadFilter to load only visible worksheets and then saves the workbook. | Explain how the VisibleSheetLoadFilter overrides StartSheet to skip hidden sheets and extend it to also ignore very hidden sheets. | Generate a unit test in C# that confirms hidden worksheets are not loaded when the custom LoadFilter is applied.

using System;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells LoadOptions with a custom LoadFilter that loads only visible worksheets. The example creates a VisibleSheetLoadFilter, applies it to LoadOptions, opens the workbook, lists the loaded sheets, and saves the result, reducing memory usage and processing time.
class Program
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string filePath = "input.xlsx";

        // Create LoadOptions and assign a custom LoadFilter that skips hidden sheets
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new VisibleSheetLoadFilter();

        // Load the workbook using the configured options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Verify that only visible worksheets were loaded
        Console.WriteLine("Loaded worksheets:");
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
        }

        // Save the workbook if further processing is required
        workbook.Save("output.xlsx");
    }

    // Custom LoadFilter implementation that loads data only for visible worksheets
    private class VisibleSheetLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load the sheet only when it is visible; otherwise skip loading its data
            if (sheet.IsVisible)
            {
                base.StartSheet(sheet);
            }
        }
    }
}
