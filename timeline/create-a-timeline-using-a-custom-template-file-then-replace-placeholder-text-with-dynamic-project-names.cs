// Title: Create a project timeline Excel workbook from a custom template and replace placeholder tags with dynamic project names using Aspose.Cells for C#
// Common Searches: C# Aspose.Cells replace {ProjectName} placeholder in Excel template | How to generate a timeline Excel file from a predefined template using Aspose.Cells | Using Aspose.Cells FindOptions to locate and replace text in a worksheet C# | Save modified Excel workbook as Xlsx with Aspose.Cells after placeholder substitution | Load custom Excel template and insert project names programmatically in C#
// Tags: Aspose.Cells find placeholder and substitute value | import Excel template with Aspose.Cells | replace worksheet placeholders using dictionary in C# | export workbook to Xlsx format using Aspose.Cells | create timeline Excel workbook programmatically | FindOptions LookAtType.Contains example

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace TimelineGenerator
{
    // The C# program loads a custom Excel template using Aspose.Cells, searches for placeholder tokens such as {ProjectName1} with FindOptions (LookAtType.Contains), replaces each token with the corresponding project name from a dictionary, ensures the output directory exists, and saves the modified workbook as ProjectTimeline.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the custom template file that contains the timeline layout
                string templatePath = @"C:\Templates\TimelineTemplate.xlsx";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Define dynamic project names to replace placeholders
                // Keys are the placeholder texts present in the template (e.g., {ProjectName1})
                // Values are the actual project names to insert
                var placeholderValues = new Dictionary<string, string>
                {
                    { "{ProjectName1}", "Alpha Initiative" },
                    { "{ProjectName2}", "Beta Expansion" },
                    { "{ProjectName3}", "Gamma Migration" }
                };

                // Iterate through each placeholder and replace it in the worksheet
                foreach (var kvp in placeholderValues)
                {
                    // Configure find options
                    FindOptions findOptions = new FindOptions
                    {
                        LookInType = LookInType.Values,
                        // Use Contains to locate the exact placeholder text; placeholders are unique
                        LookAtType = LookAtType.Contains,
                        CaseSensitive = false
                    };

                    // Start searching from the first cell
                    Cell foundCell = worksheet.Cells.Find(kvp.Key, null, findOptions);

                    while (foundCell != null)
                    {
                        // Replace the placeholder with the actual project name
                        foundCell.PutValue(kvp.Value);

                        // Continue searching from the next cell
                        foundCell = worksheet.Cells.Find(kvp.Key, foundCell, findOptions);
                    }
                }

                // Ensure the output directory exists
                string outputPath = @"C:\Outputs\ProjectTimeline.xlsx";
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to a new file
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine("Timeline generated successfully at: " + outputPath);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Failed to save the workbook:");
                    Console.WriteLine(saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the timeline:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
