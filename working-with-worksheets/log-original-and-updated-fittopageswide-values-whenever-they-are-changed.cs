// Title: How to log original and updated FitToPagesWide values for each worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, iterates every worksheet, prints the current PageSetup.FitToPagesWide, sets it to a new value, and prints the updated value. | Create a C# example that records the original FitToPagesWide setting of each sheet, changes it to 1, and logs both values to the console using Aspose.Cells.
// Common Searches: Aspose.Cells C# log FitToPagesWide before and after changing page setup | How to read and modify FitToPagesWide for each worksheet in a .xlsx file using Aspose.Cells | C# example to iterate worksheets and output original FitToPagesWide value with Aspose.Cells | Saving workbook after updating FitToPagesWide property in Aspose.Cells .NET
// Tags: Aspose.Cells page setup FitToPagesWide logging | C# iterate worksheets Aspose.Cells | modify worksheet scaling Aspose.Cells | record original page setup values .NET | save workbook after FitToPagesWide change

using Aspose.Cells;
using System;

// Loads 'input.xlsx', iterates each worksheet, writes the original PageSetup.FitToPagesWide value to the console, sets the property to 1, logs the updated value, and saves the workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Log the original FitToPagesWide value
            Console.WriteLine($"Worksheet '{sheet.Name}' original FitToPagesWide: {pageSetup.FitToPagesWide}");

            // Change the FitToPagesWide value (example: set to 1)
            pageSetup.FitToPagesWide = 1;

            // Log the updated FitToPagesWide value
            Console.WriteLine($"Worksheet '{sheet.Name}' updated FitToPagesWide: {pageSetup.FitToPagesWide}");
        }

        // Save the modified workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
