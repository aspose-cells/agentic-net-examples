using System;
using System.Collections.Generic;
using Aspose.Cells;

// Custom warning callback that stores warnings for later review
class LoadWarningCallback : IWarningCallback
{
    // List to keep all warning information received during load
    public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

    // This method is called by Aspose.Cells whenever a warning occurs
    public void Warning(WarningInfo warningInfo)
    {
        // Store the warning; we will output it after the workbook is loaded
        Warnings.Add(warningInfo);
    }
}

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Create LoadOptions and assign the custom warning callback
        LoadOptions loadOptions = new LoadOptions();
        var warningCallback = new LoadWarningCallback();
        loadOptions.WarningCallback = warningCallback;

        // Load the workbook using the constructor that accepts LoadOptions
        // Any warnings generated during loading will be captured by the callback
        Workbook workbook = new Workbook(filePath, loadOptions);

        // After initialization, output all captured warnings to the console
        if (warningCallback.Warnings.Count > 0)
        {
            Console.WriteLine("Workbook load warnings:");
            foreach (var warning in warningCallback.Warnings)
            {
                Console.WriteLine($"- Type: {warning.Type}, Description: {warning.Description}");
            }
        }
        else
        {
            Console.WriteLine("No load warnings detected.");
        }

        // The workbook is now ready for further processing...
    }
}