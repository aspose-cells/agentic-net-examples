// Title: Load a partially corrupted XLSX workbook with Aspose.Cells for .NET, capture load warnings via reflection, and save the recovered file
// AI Prompts: Open a damaged .xlsx file using Aspose.Cells LoadOptions, invoke GetWarnings through reflection, and list each warning message. | After retrieving load warnings, access the first worksheet name and output it to the console. | Save the workbook that was loaded despite corruption to a new file path, handling any save exceptions.
// Common Searches: Aspose.Cells .NET load corrupted Excel file and get warning messages | how to use reflection to call GetWarnings on a workbook in C# | recover data from a partially damaged xlsx with Aspose.Cells | save workbook after loading corrupted file using Aspose.Cells | handle load warnings when opening damaged Excel workbook in C#
// Tags: load corrupted xlsx Aspose.Cells .NET | invoke GetWarnings via reflection | recoverable content from damaged Excel workbook | save recovered workbook Aspose.Cells | handle workbook load warnings

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example checks for a partially corrupted XLSX file, loads it with Aspose.Cells using LoadOptions, uses reflection to call GetWarnings and prints each warning, accesses the first worksheet, and optionally saves the recovered workbook to a new file while handling errors.
class Program
{
    static void Main()
    {
        // Path to the partially corrupted Excel file
        string sourcePath = "corrupted.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Configure load options (format can be auto-detected or specified)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook while capturing any warnings generated during the process
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Attempt to retrieve warnings via reflection (compatible with multiple Aspose.Cells versions)
            try
            {
                MethodInfo getWarningsMethod = workbook.GetType().GetMethod("GetWarnings", BindingFlags.Instance | BindingFlags.Public);
                if (getWarningsMethod != null)
                {
                    var result = getWarningsMethod.Invoke(workbook, null) as IList<WarningInfo>;
                    if (result != null)
                    {
                        foreach (WarningInfo warning in result)
                        {
                            Console.WriteLine($"Warning: {warning.Description}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If reflection fails, continue without processing warnings
                Console.WriteLine($"Unable to retrieve warnings: {ex.Message}");
            }

            // At this point the workbook contains all recoverable content.
            // Example: access the first worksheet and output its name.
            if (workbook.Worksheets.Count > 0)
            {
                Worksheet firstSheet = workbook.Worksheets[0];
                Console.WriteLine($"First worksheet name: {firstSheet.Name}");
            }

            // Save the recovered workbook to a new file (optional)
            string outputPath = "recovered.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Recovered workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
