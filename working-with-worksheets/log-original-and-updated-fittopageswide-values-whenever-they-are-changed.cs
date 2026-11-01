// Title: C# example to log original and updated FitToPagesWide values for each worksheet using Aspose.Cells
// Common Searches: how to print original and new FitToPagesWide values for each sheet with Aspose.Cells in C# | Aspose.Cells C# log page setup FitToPagesWide before and after change | record FitToPagesWide property changes for worksheets using Aspose.Cells .NET | C# console output of FitToPagesWide modifications in Excel workbook with Aspose.Cells | detect and log FitToPagesWide differences across worksheets using Aspose.Cells
// Tags: Aspose.Cells FitToPagesWide property logging | C# worksheet page setup scaling update | detect and record page setup changes Aspose.Cells | Excel print scaling modification tracking .NET | monitor worksheet FitToPagesWide value changes

using Aspose.Cells;
using System;

// The code loads an Excel workbook, iterates through each worksheet, captures the current PageSetup.FitToPagesWide value, sets a new value when it differs, logs both the original and updated values to the console, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet to monitor FitToPagesWide changes
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Store the original value
            int originalFitToPagesWide = pageSetup.FitToPagesWide;

            // Define the new value you want to set (example: 1)
            int newFitToPagesWide = 1;

            // Only proceed if the value is actually changing
            if (originalFitToPagesWide != newFitToPagesWide)
            {
                // Log the original value
                Console.WriteLine($"Worksheet '{sheet.Name}': Original FitToPagesWide = {originalFitToPagesWide}");

                // Apply the new value
                pageSetup.FitToPagesWide = newFitToPagesWide;

                // Log the updated value
                Console.WriteLine($"Worksheet '{sheet.Name}': Updated FitToPagesWide = {pageSetup.FitToPagesWide}");
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
