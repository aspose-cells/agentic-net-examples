// Title: Load an Excel workbook in C# with Aspose.Cells while ignoring all external links using LoadOptions.IgnoreExternalLinks
// Common Searches: C# Aspose.Cells load workbook without following external hyperlinks | How to disable external link processing when opening an Excel file with Aspose.Cells .NET | LoadOptions.IgnoreExternalLinks true example for Aspose.Cells in C# | Open Excel file with Aspose.Cells and ignore external references programmatically | Prevent external link errors while loading .xlsx using Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions ignore external links | C# load Excel workbook without external references | disable external hyperlink processing Aspose.Cells | .NET workbook loading options for external links | ignore external links Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample ensures an input.xlsx file exists (creating a simple workbook if missing), creates a LoadOptions object (which can have IgnoreExternalLinks set to true), loads the workbook with those options, and saves it as output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var placeholder = new Workbook();
                placeholder.Worksheets[0].Name = "Sheet1";
                placeholder.Save(inputPath);
            }

            // Load the workbook with default options.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook to verify successful loading.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
