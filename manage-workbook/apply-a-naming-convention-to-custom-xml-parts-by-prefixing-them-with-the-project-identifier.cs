// Title: How to prefix custom XML part names with a project identifier in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Add a project ID prefix to each CustomXmlPart name in a workbook and save the changes with Aspose.Cells in C#. | Iterate through Workbook.CustomXmlParts, verify the prefix is absent, and rename parts only when needed before saving.
// Common Searches: C# Aspose.Cells rename custom XML parts with project prefix | how to add a prefix to custom XML part names in an Excel file using Aspose.Cells | iterate custom XML parts in a workbook and modify their names before saving | apply naming convention to CustomXmlParts collection Aspose.Cells .NET | prevent duplicate prefix when renaming custom XML parts in Excel workbook
// Tags: customxmlparts rename Aspose.Cells C# | prefix workbook custom xml parts | apply naming convention to Excel custom XML parts | Aspose.Cells custom XML part management | avoid duplicate prefix custom xml parts

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, loops through its CustomXmlParts collection, and prepends a specified project identifier to each part's Name only if the prefix is not already present, then saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Project identifier to be used as a prefix
        string projectId = "Proj123";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate over all custom XML parts and rename them
            foreach (dynamic xmlPart in workbook.CustomXmlParts)
            {
                string originalName = xmlPart.Name;

                // Apply the prefix only if it hasn't been added already
                if (!originalName.StartsWith(projectId + "_"))
                {
                    xmlPart.Name = $"{projectId}_{originalName}";
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated names (lifecycle: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
