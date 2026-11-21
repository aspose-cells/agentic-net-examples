// Title: Use a loop to assign custom GlobalizationSettings (subtotal label language) to each worksheet via custom properties in Aspose.Cells for .NET
// Common Searches: how to store worksheet-specific subtotal label language in Aspose.Cells using C# | apply different globalization settings to multiple Excel sheets with a loop in Aspose.Cells | set custom property for culture info on each worksheet in Aspose.Cells .NET | assign unique subtotal label language per sheet when SubtotalLabelLanguage property is unavailable | loop through worksheets to add localization data with Aspose.Cells API
// Tags: loop assign worksheet globalization custom property | custom subtotal label language Aspose.Cells | worksheet culture info via custom properties C# | multiple sheet localization Aspose.Cells | apply per-sheet globalization settings .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, adds three worksheets, defines an array of culture codes (en-US, fr-FR, de-DE), and iterates through the sheets. For each worksheet it stores the corresponding culture code in a custom property named "SubtotalLabelLanguage" because the direct SubtotalLabelLanguage API is not available. The workbook is then saved as CustomGlobalizationSettings.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets that will receive custom globalization settings
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Define the subtotal label languages for each worksheet
            string[] subtotalLabelLanguages = { "en-US", "fr-FR", "de-DE" };

            // Loop through the worksheets and store the language as a custom property
            for (int i = 0; i < workbook.Worksheets.Count && i < subtotalLabelLanguages.Length; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Store the language information in a custom property (since GlobalizationSettings.SubtotalLabelLanguage is not available)
                sheet.CustomProperties.Add("SubtotalLabelLanguage", subtotalLabelLanguages[i]);
            }

            // Define output file path
            string outputPath = "CustomGlobalizationSettings.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
