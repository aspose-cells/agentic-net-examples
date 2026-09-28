// Title: How to set Worksheet TabId to zero and capture the resulting exception using Aspose.Cells for .NET
// AI Prompts: Assign a TabId value of 0 to every worksheet in a workbook and record any exception that is thrown. | Create code that logs the exception message when setting Worksheet.TabId to zero fails with Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet TabId to 0 and handle exception | C# catch exception when assigning zero TabId to worksheets in Aspose.Cells | Logging errors for invalid TabId value in Aspose.Cells workbook | Why does setting Worksheet.TabId to zero throw an exception in Aspose.Cells .NET
// Tags: worksheet TabId zero exception Aspose.Cells | Aspose.Cells set TabId property error handling | C# Aspose.Cells workbook worksheet TabId assignment | exception logging for Aspose.Cells worksheet TabId | Aspose.Cells .NET worksheet TabId validation

using System;
using Aspose.Cells;

// The example creates a new workbook, removes the default sheet, adds two uniquely named worksheets, then iterates through each sheet attempting to set its TabId to zero. The expected exception is caught and logged, and the workbook is saved as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Remove the default worksheet to avoid duplicate names
            workbook.Worksheets.Clear();

            // Add worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");

            // Iterate through all worksheets and attempt to set TabId to zero
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // This operation is expected to throw an exception
                    sheet.TabId = 0;
                    Console.WriteLine($"Worksheet '{sheet.Name}': TabId set to 0 successfully.");
                }
                catch (Exception ex)
                {
                    // Log the expected exception
                    Console.WriteLine($"Worksheet '{sheet.Name}': Exception caught - {ex.Message}");
                }
            }

            // Save the workbook
            workbook.Save("Result.xlsx");
            Console.WriteLine("Workbook saved as 'Result.xlsx'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
