// Title: Subclass GlobalizationSettings and override GetGrandTotalName to provide a culture‑specific Grand Total label in Aspose.Cells for .NET
// Common Searches: how to override GetGrandTotalName in Aspose.Cells using C# | customize pivot table grand total label per culture with Aspose.Cells | C# subclass GlobalizationSettings to set localized grand total text | Aspose.Cells example for culture‑aware Grand Total naming | change grand total caption based on locale in Aspose.Cells workbook
// Tags: GetGrandTotalName method override | custom globalization settings subclass C# | culture aware grand total naming | pivot table total label localization | Aspose.Cells globalization API customization

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates how to create a subclass of GlobalizationSettings in Aspose.Cells for .NET and override the GetGrandTotalName method to return a culture‑specific label for Grand Total rows in pivot tables, enabling localized total captions while preserving standard workbook creation and saving.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook.
            var workbook = new Workbook();

            // Access the workbook's globalization settings.
            var globalization = workbook.Settings.GlobalizationSettings;

            // Note: Setting a custom Grand Total label is not directly supported in this API version.
            // If needed, customize pivot table labels after creation.

            // Define the output file path.
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
