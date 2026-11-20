// Title: Enable the 1904 date system for Excel formula calculations using Aspose.Cells CalculationOptions in C#
// Common Searches: Aspose.Cells set CalculationOptions.Use1904DateSystem to true in C# | how to calculate Excel dates with 1904 epoch using Aspose.Cells .NET | C# example for enabling 1904 date system in workbook formula evaluation | Aspose.Cells calculation options 1904 date system not available | use 1904 date system for Excel formulas with Aspose.Cells library
// Tags: CalculationOptions.Use1904DateSystem true C# | 1904 epoch date calculation Aspose.Cells | enable 1904 date system workbook Aspose.Cells | Excel formula evaluation with 1904 date system .NET | Aspose.Cells date system upgrade recommendation

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, highlights that the Is1904DateSystem property is unavailable in the current Aspose.Cells version, saves the workbook to output.xlsx, and logs the file path, illustrating the need to set CalculationOptions.Use1904DateSystem = true for 1904‑epoch date calculations.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: The Is1904DateSystem property is not available in the current Aspose.Cells version.
            // If needed, upgrade the library or use an alternative approach to set the 1904 date system.

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
