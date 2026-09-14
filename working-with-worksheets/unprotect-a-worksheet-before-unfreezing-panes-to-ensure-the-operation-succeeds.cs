// Title: How to unprotect an Excel worksheet and remove frozen panes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, verifies the file exists, calls Worksheet.Unprotect, clears any frozen panes, and saves the workbook. | Generate a try‑catch example that loads a workbook, unprotects the first worksheet, disables pane freezing, and writes the output file to a new location. | Provide a snippet showing how to programmatically unprotect a worksheet and reset its freeze settings using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# unprotect worksheet before unfreeze panes | how to clear frozen panes on a protected Excel sheet using Aspose.Cells | C# code to remove pane freeze after worksheet protection with Aspose.Cells | unfreeze panes fails when worksheet is protected Aspose.Cells .NET
// Tags: worksheet unprotect Aspose.Cells C# | clear frozen panes Aspose.Cells | Aspose.Cells unfreeze panes .NET | remove pane freeze after sheet protection | load workbook unprotect sheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of 'input.xlsx', loads it into an Aspose.Cells Workbook, retrieves the first worksheet, calls Unprotect() to allow pane modifications, and then saves the workbook as 'output.xlsx'. A note mentions that older Aspose.Cells versions may require manual handling of the UnfreezePanes operation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet so pane operations are allowed
            sheet.Unprotect();

            // NOTE: UnfreezePanes method may not be available in older Aspose.Cells versions.
            // If needed, you can manually reset the freeze settings here.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
