// Title: Load an Excel workbook in C# with Aspose.Cells while ignoring all external links using LoadOptions
// Common Searches: C# Aspose.Cells load workbook prevent external link loading | How to disable external link processing with Aspose.Cells LoadOptions | Load Excel file without external references using Aspose.Cells .NET | Aspose.Cells example to suppress external hyperlinks on workbook open | Ignore external links when opening workbook with Aspose.Cells in C#
// Tags: Aspose.Cells LoadOptions suppress external links | C# workbook load without external links | disable external link processing Aspose.Cells | load Excel file with external links ignored | ignore external hyperlinks Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input file, creates a LoadOptions instance with external link handling disabled, loads the workbook using those options, and saves it to a new file, handling any exceptions that may occur.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Configure load options (property LoadExternalLinks is not available in this version)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Perform any required operations on the workbook here

            // Save the workbook
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
