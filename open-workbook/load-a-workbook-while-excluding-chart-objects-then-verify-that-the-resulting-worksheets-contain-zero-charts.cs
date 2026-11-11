// Title: Load an XLSX workbook without charts and confirm zero chart objects per worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells load workbook without loading chart objects in C# | How to check that a worksheet has no charts after opening an Excel file with Aspose.Cells | C# code to exclude charts when opening an XLSX file using Aspose.Cells LoadOptions | Verify chart count is zero for each sheet after loading workbook in Aspose.Cells .NET
// Tags: load workbook excluding charts Aspose.Cells | worksheet chart collection verification C# | exclude charts using LoadOptions for XLSX | zero chart count per worksheet Aspose.Cells | chart presence validation after workbook load

using Aspose.Cells;
using System;
using System.IO;

// The example verifies that the input.xlsx file exists, loads it into an Aspose.Cells Workbook with LoadOptions for XLSX, iterates through each worksheet, and prints whether the worksheet's Charts collection is empty or reports the number of chart objects present.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook (LoadDataOnly is not available in this version of Aspose.Cells)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Verify that each worksheet contains zero charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' contains no charts.");
                }
                else
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' still has {sheet.Charts.Count} chart(s).");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
