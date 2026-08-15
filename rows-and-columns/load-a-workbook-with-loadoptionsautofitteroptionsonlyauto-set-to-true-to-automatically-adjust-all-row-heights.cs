// Title: C# – Load Excel with Aspose.Cells AutoFitterOptions.OnlyAuto to auto‑fit rows
// Description: Shows how to configure LoadOptions with AutoFitterOptions.OnlyAuto = true so that Aspose.Cells automatically adjusts the height of rows that are not manually sized when a workbook is opened.
// Keywords: Aspose.Cells | AutoFitterOptions | OnlyAuto | LoadOptions | C# | auto‑fit rows | Excel row height | load workbook | adjust row height | Aspose.Cells .NET example
// Common Searches: Aspose.Cells load workbook auto‑fit rows C# | AutoFitterOptions OnlyAuto example | How to auto‑adjust row heights on load with Aspose.Cells | C# load Excel file with automatic row height | Aspose.Cells LoadOptions row height settings
// Developer Intent: Automatically adjust non‑custom row heights while loading an Excel file using Aspose.Cells.
// Use Cases: Open existing spreadsheets and ensure all default rows fit their content without altering manually sized rows. | Validate row height after loading to confirm auto‑fit behavior. | Batch‑process multiple workbooks, applying OnlyAuto auto‑fit during load before further data manipulation.
// AI Prompts: Generate C# code that loads an Excel workbook with LoadOptions.AutoFitterOptions.OnlyAuto set to true and saves it. | Explain the effect of AutoFitterOptions.OnlyAuto on rows with custom heights versus default rows. | Combine AutoFitterOptions.OnlyAuto with other LoadOptions features, such as loading specific worksheets or preserving formulas.

using System;
using Aspose.Cells;

// Shows how to configure LoadOptions with AutoFitterOptions.OnlyAuto = true so that Aspose.Cells automatically adjusts the height of rows that are not manually sized when a workbook is opened.
class Program
{
    static void Main()
    {
        // Create AutoFitterOptions and enable OnlyAuto to auto‑fit rows that are not custom‑sized
        AutoFitterOptions autoFitOptions = new AutoFitterOptions
        {
            OnlyAuto = true
        };

        // Assign the options to LoadOptions
        LoadOptions loadOptions = new LoadOptions
        {
            AutoFitterOptions = autoFitOptions
        };

        // Load the workbook with the specified load options
        // This will automatically adjust all row heights according to the OnlyAuto setting
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Example: output the height of the first row after loading
        double firstRowHeight = workbook.Worksheets[0].Cells.GetRowHeight(0);
        Console.WriteLine("First row height after auto‑fit: " + firstRowHeight);

        // Save the workbook (using the standard save rule)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
