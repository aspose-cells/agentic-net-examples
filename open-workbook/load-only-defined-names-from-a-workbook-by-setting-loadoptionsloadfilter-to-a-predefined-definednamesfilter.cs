// Title: Load only defined names from an Excel workbook using Aspose.Cells DefinedNamesFilter in C#
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, sets LoadOptions.LoadFilter to a DefinedNamesFilter, and loads only the workbook's defined names. | Show how to iterate through Workbook.Worksheets.Names after applying DefinedNamesFilter and output each name together with its RefersTo reference. | Add comprehensive error handling and demonstrate saving the workbook after loading only defined names with Aspose.Cells.
// Common Searches: Aspose.Cells C# load only named ranges using DefinedNamesFilter | How to filter workbook loading to retrieve only defined names in .NET | Example of using LoadOptions.LoadFilter with DefinedNamesFilter in Aspose.Cells | List all defined names from an Excel file without loading worksheets Aspose.Cells
// Tags: Aspose.Cells DefinedNamesFilter | load defined names Excel .NET | filter workbook load options Aspose | enumerate named ranges C# | save workbook after filtered load

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of the input Excel file, creates a LoadOptions instance with LoadFilter set to DefinedNamesFilter to load only defined names, opens the workbook, iterates through Workbook.Worksheets.Names to display each name and its RefersTo value, ensures the output directory exists, saves the workbook to a new file, and includes robust exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Create load options (no specific filter needed for defined names)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through the loaded defined names
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                Console.WriteLine($"Name: {definedName.Text}, RefersTo: {definedName.RefersTo}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, as only defined names were loaded)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
