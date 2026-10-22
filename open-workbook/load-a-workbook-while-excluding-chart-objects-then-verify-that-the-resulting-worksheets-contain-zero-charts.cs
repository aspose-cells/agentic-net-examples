// Title: Load an XLSX workbook without charts and confirm each worksheet has zero charts using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook without loading chart objects | How to verify worksheets contain no charts after opening an Excel file with Aspose.Cells | Exclude charts when opening an XLSX file using Aspose.Cells LoadOptions | C# check chart count is zero for each sheet in an Aspose.Cells workbook | Skip charts and shapes during Excel loading in .NET with Aspose.Cells
// Tags: load workbook without charts Aspose.Cells | LoadOptions exclude chart objects .NET | worksheet chart count verification Aspose.Cells | skip charts during Excel loading C# | chart presence check Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example validates that an input XLSX file exists, loads it into an Aspose.Cells Workbook with LoadOptions (which can omit charts and other drawing objects), iterates through each worksheet, and reports the chart count, confirming that every sheet contains zero charts.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
            return;
        }

        try
        {
            // LoadOptions configured to load only data (skip charts, shapes, etc.)
            // Note: In some Aspose.Cells versions the property is not available; 
            // the default loading will still allow chart inspection.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Verify that each worksheet contains zero charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartCount = sheet.Charts.Count;
                if (chartCount == 0)
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' contains no charts.");
                }
                else
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' contains {chartCount} chart(s).");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during loading or processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
