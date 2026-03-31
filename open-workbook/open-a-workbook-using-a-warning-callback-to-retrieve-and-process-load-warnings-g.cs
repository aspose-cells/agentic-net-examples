using System;
using Aspose.Cells;

// Custom warning callback implementation to capture load warnings
class CustomWarningCallback : IWarningCallback
{
    public void Warning(WarningInfo warningInfo)
    {
        // Output warning type and description to the console
        Console.WriteLine($"Warning Type: {warningInfo.Type}, Description: {warningInfo.Description}");
    }
}

class Program
{
    static void Main()
    {
        // Initialize LoadOptions and assign the custom warning callback
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.WarningCallback = new CustomWarningCallback();

        // Load the workbook with the specified options; any load warnings will trigger the callback
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Example usage after loading: display the number of worksheets
        Console.WriteLine($"Workbook loaded. Worksheet count: {workbook.Worksheets.Count}");
    }
}