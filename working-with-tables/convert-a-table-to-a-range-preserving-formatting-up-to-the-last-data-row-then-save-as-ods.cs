// Title: Convert an Excel ListObject to a regular range while preserving formatting and save as ODS using Aspose.Cells for .NET
// AI Prompts: Locate the first ListObject in a worksheet, call ConvertToRange to keep all formatting, then save the workbook as an ODS file with Aspose.Cells. | Write C# code that checks for a table, converts it to a normal range without losing styles, and exports the result to OpenDocument Spreadsheet format.
// Common Searches: Aspose.Cells convert Excel table to range preserving formatting .NET | C# save workbook as ODS after removing table definition | How to export a worksheet with converted ListObject to ODS using Aspose.Cells
// Tags: Aspose.Cells ListObject conversion to regular range | retain cell styles when converting Excel tables | save workbook as OpenDocument Spreadsheet Aspose.Cells | delete Excel table definition via Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads input.xlsx, checks for a ListObject in the first worksheet, converts the table to a normal range while preserving all formatting, and saves the workbook as output.ods.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Retrieve the first table
                ListObject table = sheet.ListObjects[0];

                // Convert the table to a normal range while preserving all formatting
                // The method removes the table definition and leaves a regular range
                table.ConvertToRange();
            }

            // Save the modified workbook in ODS format (use the non‑obsolete enum value)
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
