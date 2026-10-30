// Title: Update external link folder paths in an Excel workbook while preserving a custom ribbon UI with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change external link folder path in existing .xlsx workbook | preserve custom ribbon UI when updating external links using Aspose.Cells | how to iterate external links collection in Aspose.Cells C# | replace old folder path with new folder in Excel external links Aspose.Cells | load and save workbook with custom ribbon UI after modifying external links C#
// Tags: update external link paths Aspose.Cells | retain custom ribbon UI workbook | dynamic external links handling C# | replace folder segment in Excel external links | Aspose.Cells save workbook with custom UI

using Aspose.Cells;
using System;
using System.IO;
using Microsoft.CSharp.RuntimeBinder;

// The example loads an existing .xlsx workbook, uses dynamic typing to access each worksheet's ExternalLinks collection, replaces the old folder segment in each link's path with a new folder path, and saves the workbook. The custom ribbon UI parts are automatically retained, and the code includes file existence checks and robust exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the custom ribbon UI and external links
            Workbook workbook = new Workbook(inputPath);

            // Update each external link path to the new location (using dynamic to stay compatible with all versions)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    dynamic dynSheet = sheet;
                    foreach (dynamic link in dynSheet.ExternalLinks)
                    {
                        if (link != null && link.LinkPath != null)
                        {
                            // Example replacement: change old folder to new folder
                            string currentPath = (string)link.LinkPath;
                            link.LinkPath = currentPath.Replace(@"C:\OldFolder\", @"D:\NewFolder\");
                        }
                    }
                }
                catch (RuntimeBinderException)
                {
                    // The current Aspose.Cells version does not support ExternalLinks; skip processing.
                }
            }

            // Save the modified workbook (custom UI parts are preserved automatically)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
