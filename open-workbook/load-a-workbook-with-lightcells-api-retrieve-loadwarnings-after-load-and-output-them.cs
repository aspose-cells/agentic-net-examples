// Title: Capture Workbook Load Warnings with Aspose.Cells LightCells API (C#)
// Description: Learn how to implement a custom IWarningCallback, assign it to LoadOptions, and use the LightCells API to load an Excel file while collecting and displaying all load warnings (type and description) in C#.
// Keywords: Aspose.Cells LightCells C# | Workbook load warnings | IWarningCallback example | LoadOptions WarningCallback | Aspose.Cells warning handling | Excel file load diagnostics | C# Aspose.Cells tutorial
// Common Searches: Aspose.Cells capture load warnings C# | How to use IWarningCallback with LightCells | Retrieve warning messages after loading workbook Aspose | LoadOptions WarningCallback example | Log Excel load warnings using Aspose.Cells
// Developer Intent: The developer needs to capture and display any warnings generated while opening an Excel workbook with the LightCells API.
// Use Cases: Log load warnings to a file for troubleshooting malformed spreadsheets. | Validate workbook compatibility by examining warning types after import. | Show a summary of load issues in a UI after a user opens a spreadsheet.
// AI Prompts: Create a C# method that loads an Excel workbook with Aspose.Cells LightCells API and returns a list of warning messages. | Show how to filter specific warning types in a custom IWarningCallback during workbook load. | Demonstrate integrating the warning callback with asynchronous workbook loading in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Custom warning callback that stores all warnings received during loading
// Learn how to implement a custom IWarningCallback, assign it to LoadOptions, and use the LightCells API to load an Excel file while collecting and displaying all load warnings (type and description) in C#.
class CustomWarningCallback : IWarningCallback
{
    // List to keep warning information
    public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

    // This method is called by Aspose.Cells when a warning occurs
    public void Warning(WarningInfo warningInfo)
    {
        Warnings.Add(warningInfo);
    }
}

class Program
{
    static void Main()
    {
        // Path to the workbook that will be loaded
        string filePath = "input.xlsx";

        // Create LoadOptions and assign the custom warning callback
        LoadOptions loadOptions = new LoadOptions();
        CustomWarningCallback warningCallback = new CustomWarningCallback();
        loadOptions.WarningCallback = warningCallback;

        // Load the workbook using the LightCells API (via LoadOptions)
        Workbook workbook = new Workbook(filePath, loadOptions);

        // After loading, output all collected warnings
        Console.WriteLine("Load warnings:");
        foreach (var warning in warningCallback.Warnings)
        {
            Console.WriteLine($"- Type: {warning.WarningType}, Description: {warning.Description}");
        }
    }
}
