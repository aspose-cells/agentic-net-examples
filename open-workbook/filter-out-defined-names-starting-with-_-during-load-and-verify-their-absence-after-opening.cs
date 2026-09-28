// Title: Remove underscore‑prefixed defined names from an Excel workbook using Aspose.Cells in C# and verify their removal before saving
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates through the NameCollection, deletes every defined name whose text starts with an underscore, confirms the collection no longer contains such names, and then saves the workbook. | Show how to filter named ranges by a leading '_' character, remove them from the worksheet's Names collection, and programmatically check that the removal succeeded before exporting the file.
// Common Searches: aspnet c# remove named ranges that begin with '_' using Aspose.Cells | how to delete underscore prefixed defined names when loading an Excel file with Aspose.Cells | verify that no named ranges start with underscore after opening workbook Aspose.Cells C# | Aspose.Cells filter NameCollection by prefix and save cleaned workbook | C# code sample for cleaning up defined names in Excel with Aspose.Cells
// Tags: aspocells remove underscore defined names | c# filter NameCollection by prefix | aspocells delete named ranges excel | c# verify named range removal aspocells | aspocells load workbook clean defined names

using Aspose.Cells;
using System;
using System.IO;
using System.Linq;

// // Loads an Excel file, removes any defined names that start with an underscore, checks that none remain, and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var loadOptions = new LoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Access the collection of defined names
            NameCollection definedNames = workbook.Worksheets.Names;

            // Identify defined names that start with '_' and collect their texts
            var namesToRemove = definedNames
                .Cast<Name>()
                .Where(n => n.Text.StartsWith("_", StringComparison.Ordinal))
                .Select(n => n.Text)
                .ToList();

            // Remove the identified defined names from the collection
            foreach (string name in namesToRemove)
            {
                definedNames.Remove(name);
            }

            // Verify that no defined name beginning with '_' remains
            bool underscoreNamesExist = definedNames
                .Cast<Name>()
                .Any(n => n.Text.StartsWith("_", StringComparison.Ordinal));

            Console.WriteLine(underscoreNamesExist
                ? "Underscore defined names still exist."
                : "All underscore defined names have been removed.");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
