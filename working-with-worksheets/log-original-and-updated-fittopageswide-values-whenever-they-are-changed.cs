// Title: Log original and new FitToPagesWide values for each worksheet using Aspose.Cells for .NET
// Common Searches: c# log worksheet FitToPagesWide before and after change using Aspose.Cells | asp.net detect and record FitToPagesWide modifications in Excel workbook | how to track FitToPagesWide setting per worksheet with Aspose.Cells .NET | output console message when FitToPagesWide is updated in C# Aspose.Cells | record original and updated page setup FitToPagesWide values in Excel file
// Tags: FitToPagesWide audit Aspose.Cells | worksheet page setup value tracking C# | Aspose.Cells update FitToPagesWide setting | Excel workbook page setup change detection | console output for page setup changes .NET

using Aspose.Cells;
using System;

// The example opens an Excel workbook, loops through each worksheet, reads the current PageSetup.FitToPagesWide, compares it to a target value (1), writes a console line showing the worksheet name and the before‑and‑after values when they differ, applies the new setting, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the PageSetup object for the current worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Capture the original FitToPagesWide value
            int originalFitToPagesWide = pageSetup.FitToPagesWide;

            // Define the new FitToPagesWide value (example: set to 1)
            int newFitToPagesWide = 1;

            // Update only if the value is actually changing
            if (originalFitToPagesWide != newFitToPagesWide)
            {
                // Log the change
                Console.WriteLine($"Worksheet '{sheet.Name}': FitToPagesWide changed from {originalFitToPagesWide} to {newFitToPagesWide}");

                // Apply the new value
                pageSetup.FitToPagesWide = newFitToPagesWide;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
