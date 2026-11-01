// Title: How to set worksheet TabId to zero for all sheets and handle the resulting exception using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set worksheet TabId to 0 and catch exception in C# | C# loop through workbook worksheets and handle TabId assignment error with Aspose.Cells | How to log exception when setting TabId property to zero in Aspose.Cells | Saving workbook after attempting to set TabId to zero for each sheet using Aspose.Cells | Why does assigning zero to worksheet TabId throw an exception in Aspose.Cells
// Tags: worksheet TabId modification Aspose.Cells | exception handling for TabId property C# | iterate workbook worksheets Aspose.Cells | log TabId assignment errors Aspose.Cells | save workbook after property changes Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel file, iterates over every worksheet, attempts to set each sheet's TabId to zero (which triggers an exception), catches and logs the exception message, and finally saves the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (using the provided load rule)
        var workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                // Attempt to set TabId to zero (expected to throw)
                sheet.TabId = 0;
                Console.WriteLine($"Worksheet '{sheet.Name}': TabId set to 0 successfully.");
            }
            catch (Exception ex)
            {
                // Log the expected exception
                Console.WriteLine($"Worksheet '{sheet.Name}': Exception when setting TabId - {ex.Message}");
            }
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("output.xlsx");
    }
}
