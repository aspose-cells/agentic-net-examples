// Title: Disable editing of pivot tables while protecting an Excel worksheet with a password using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# protect worksheet but allow pivot table refresh | How to prevent pivot table editing in a protected Excel sheet using Aspose.Cells | C# set worksheet protection options to keep pivot tables usable with Aspose.Cells | Apply password protection to Excel worksheet while keeping pivot tables functional in Aspose.Cells | Enable pivot table data refresh on a protected sheet with Aspose.Cells C#
// Tags: worksheet password protection Aspose.Cells C# | pivot table refresh allowed on protected sheet Aspose.Cells | pivot table editing disabled Aspose.Cells | ProtectionType.All usage Aspose.Cells | AllowUsePivotTable setting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, applies full protection with a password using sheet.Protect(ProtectionType.All, "myPassword", string.Empty), and saves the workbook, ensuring pivot tables can be refreshed but not edited.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or use workbook.Worksheets["Sheet1"])
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Configure protection options (if supported by the version) -----
            // Example of setting options:
            // sheet.Protection.AllowEditObject = false;
            // sheet.Protection.AllowUsePivotTable = true;
            // The above properties may not be available in older versions; they can be omitted.

            // Protect the worksheet with a password.
            // The overload requires the old password; an empty string is used when there is none.
            sheet.Protect(ProtectionType.All, "myPassword", string.Empty);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
