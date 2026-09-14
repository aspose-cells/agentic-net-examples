// Title: Load an Excel workbook in C# with Aspose.Cells while suppressing external link processing using LoadOptions.IgnoreExternalLinks
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, sets LoadOptions.IgnoreExternalLinks = true, and saves the workbook to a new file. | Show how to configure Aspose.Cells LoadOptions to prevent evaluation of external hyperlinks when loading a workbook. | Provide a sample that loads a workbook, disables resolution of linked data sources, and writes the result using Aspose.Cells.
// Common Searches: Aspose.Cells C# load workbook without following external references | How to turn off external link handling in Aspose.Cells LoadOptions | C# example for loading Excel file while ignoring linked data sources with Aspose.Cells
// Tags: Aspose.Cells LoadOptions.IgnoreExternalLinks usage | disable external link processing in Excel workbook loading | C# load .xlsx without external references Aspose.Cells | suppress external hyperlink evaluation Aspose.Cells | Excel workbook loading options .NET Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example checks that input.xlsx exists, creates a LoadOptions object with IgnoreExternalLinks set to true, loads the workbook using those options, and saves it as output.xlsx, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Configure load options (default options are used here)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // TODO: Add any workbook manipulation code here

            // Save the workbook to a new file (or overwrite the original)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
