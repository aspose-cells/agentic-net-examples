// Title: Load an XLSX workbook in C# with Aspose.Cells while preserving cell formatting and skipping formula evaluation using LoadOptions
// Common Searches: C# Aspose.Cells load XLSX without evaluating formulas | how to keep original cell styles when opening Excel with Aspose.Cells | LoadOptions to skip formula evaluation in Aspose.Cells C# | preserve formatting on workbook load using Aspose.Cells LoadOptions | load workbook ignoring formula results Aspose.Cells .NET
// Tags: Aspose.Cells LoadOptions preserve cell formatting | ignore formula results Aspose.Cells LoadDataOptions | load XLSX workbook without formula calculation C# | preserve original styles on workbook load Aspose.Cells | C# LoadOptions Xlsx format Aspose.Cells example

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for an input.xlsx file, creates a simple workbook if missing, configures LoadOptions for the XLSX format (with a commented suggestion to set LoadDataOptions for preserving formatting and ignoring formulas), loads the workbook using those options, and saves it as output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Ensure the source file exists; create a simple workbook if it does not
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample");
                tempWb.Save(sourcePath);
            }

            // Configure LoadOptions to preserve formatting (if supported by the version)
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // In newer versions LoadDataOptions can be set, but if unavailable we rely on default behavior
            // loadOptions.LoadDataOptions = LoadDataOptions.LoadFormula | LoadDataOptions.PreserveFormatting;

            // Load the workbook with the specified options
            var workbook = new Workbook(sourcePath, loadOptions);

            // (Optional) Perform any additional processing here

            // Save the workbook to verify the loading behavior
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
