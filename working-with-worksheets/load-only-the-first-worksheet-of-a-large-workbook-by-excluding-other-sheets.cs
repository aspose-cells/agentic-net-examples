// Title: C# – Load Only the First Worksheet from a Large Excel Workbook Using Aspose.Cells LoadFilter
// Description: Demonstrates how to create a custom LoadFilter that loads just the first sheet (index 0) of a massive Excel file, reducing memory usage. The filter is applied via LoadOptions, the workbook is opened, and the single‑sheet result can be saved if needed.
// Keywords: Aspose.Cells LoadFilter C# | load first worksheet only | load specific sheet Aspose.Cells | memory‑efficient Excel loading .NET | exclude worksheets on load | large workbook performance Aspose | SheetsInLoadingOrder example
// Common Searches: Aspose.Cells load only first sheet | C# LoadFilter to read single worksheet | how to avoid loading all sheets in Aspose.Cells | reduce memory when opening big Excel file .NET | load specific worksheet by index Aspose
// Developer Intent: Open an Excel workbook while loading only the first worksheet into memory.
// Use Cases: Minimize RAM consumption for applications that process huge workbooks but need data from only the first sheet. | Create a lightweight copy of the initial worksheet for quick analysis or transformation. | Speed up application startup when the first sheet contains all required information.
// AI Prompts: Provide a C# example that loads a worksheet by its name using Aspose.Cells LoadFilter. | Show how to modify the LoadFirstSheetFilter to load multiple sheets, such as indexes 0 and 2. | Explain how to combine LoadFilter with LoadDataFilterOptions.None to retrieve only sheet metadata without cell values.

using System;
using Aspose.Cells;

// Demonstrates how to create a custom LoadFilter that loads just the first sheet (index 0) of a massive Excel file, reducing memory usage. The filter is applied via LoadOptions, the workbook is opened, and the single‑sheet result can be saved if needed.
class LoadFirstSheetFilter : LoadFilter
{
    // Load all data for the sheets that are loaded
    public LoadFirstSheetFilter() : base(LoadDataFilterOptions.All) { }

    // Specify that only the sheet at index 0 should be loaded
    public override int[] SheetsInLoadingOrder => new int[] { 0 };
}

class Program
{
    static void Main()
    {
        // Set up load options with the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFirstSheetFilter();

        // Load the workbook; only the first worksheet will be loaded
        Workbook workbook = new Workbook("largeWorkbook.xlsx", loadOptions);

        // Verify that only one worksheet is present
        Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);
        Console.WriteLine("First sheet name: " + workbook.Worksheets[0].Name);

        // Save the workbook containing only the first sheet (optional)
        workbook.Save("firstSheetOnly.xlsx");
    }
}
