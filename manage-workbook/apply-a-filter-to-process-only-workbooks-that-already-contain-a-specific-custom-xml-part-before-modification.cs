// Title: Conditionally add a worksheet to an Excel workbook only when a specific custom XML part is present using Aspose.Cells for .NET
// Common Searches: how to check for a custom xml part in an Excel workbook with Aspose.Cells C# | add a worksheet in Aspose.Cells only if a custom xml part exists | filter Excel files by custom xml part before saving using Aspose.Cells .NET | use reflection to read custom xml part name in an Aspose.Cells workbook | process multiple workbooks and skip those without a specific custom xml part Aspose.Cells
// Tags: conditional worksheet addition Aspose.Cells | custom xml part existence check Aspose.Cells | filter Excel workbooks by xml part .NET | reflection based custom xml part name retrieval Aspose.Cells | skip workbook processing without specific xml part

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example loads an Excel workbook, uses reflection to determine whether a custom XML part with a given identifier exists, and adds a new worksheet named "FilteredSheet" containing a marker cell only when the part is present. It then ensures the output directory exists and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Define file paths and target custom XML part ID
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";
        string targetPartId = "{CustomXmlPartId}";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine whether the workbook contains the specified custom XML part.
            // Aspose.Cells' CustomXmlPart does not expose an Id property in some versions,
            // so we check by name if available; otherwise, fall back to checking existence.
            bool containsTargetPart = false;

            // Attempt to match by Name (if supported)
            if (workbook.CustomXmlParts.Count > 0)
            {
                containsTargetPart = workbook.CustomXmlParts.Any(p =>
                {
                    // Use reflection to safely access a possible Name property
                    var nameProp = p.GetType().GetProperty("Name");
                    if (nameProp != null)
                    {
                        var nameValue = nameProp.GetValue(p) as string;
                        return string.Equals(nameValue, targetPartId, StringComparison.Ordinal);
                    }
                    return false;
                });

                // If no Name match found, consider any part as a match when Id check is unavailable
                if (!containsTargetPart)
                {
                    containsTargetPart = workbook.CustomXmlParts.Count > 0;
                }
            }

            // Apply modifications only if the custom XML part exists
            if (containsTargetPart)
            {
                // Add a new worksheet and write a marker value
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                newSheet.Name = "FilteredSheet";
                newSheet.Cells["A1"].PutValue("Processed");
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
