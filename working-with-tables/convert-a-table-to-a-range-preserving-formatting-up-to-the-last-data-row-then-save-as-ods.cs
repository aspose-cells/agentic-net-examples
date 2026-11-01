// Title: Convert an Excel table to a regular range while keeping formatting and save as ODS with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert ListObject to range preserving cell styles | How to export an Excel worksheet containing a table to ODS format using Aspose.Cells | C# code to change an Excel table into a normal range before saving as ODS | Preserve table formatting when converting Excel to ODS with Aspose.Cells .NET | Convert first ListObject to range and save workbook as ODS using Aspose.Cells
// Tags: ListObject conversion to range Aspose.Cells | ODS export with formatting retention | Aspose.Cells save workbook as ODS | Excel table to normal range C# | retain cell styles during ODS conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an XLSX workbook, checks for a ListObject on the first worksheet, converts the first table to a normal range while preserving its formatting, and then saves the workbook as an ODS file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.ods";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Check if the worksheet contains any tables (ListObjects)
            if (sheet.ListObjects.Count > 0)
            {
                // Get the first table
                ListObject table = sheet.ListObjects[0];

                // Convert the table to a normal range while preserving formatting
                table.ConvertToRange();
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Save the workbook in ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
