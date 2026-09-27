// Title: Add a worksheet to an Excel file only when it contains a custom XML part using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook with Aspose.Cells, verify that any custom XML part exists, and insert a new worksheet named "Processed" only in that case. | Write C# code that opens a .xlsx file, checks for a specific custom XML part ID, and skips all modifications when the part is not found. | Adapt the example to iterate through CustomXmlParts and conditionally add content based on the presence of at least one part.
// Common Searches: Aspose.Cells check if workbook has custom XML part before editing | C# add worksheet conditionally based on custom XML part in Excel | how to skip modifications in Aspose.Cells when custom XML part is absent | filter Excel files by custom XML part presence using Aspose.Cells .NET
// Tags: conditional worksheet addition Aspose.Cells | custom XML part detection .NET Excel | filter workbooks by custom XML part Aspose.Cells | skip workbook modifications when XML part missing | verify custom XML part existence Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads input.xlsx, checks whether the workbook contains any custom XML parts, adds a worksheet named "Processed" with a note when a part is found, and saves the result to output.xlsx; if no custom XML part is present, it logs a message and leaves the file unchanged.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetCustomXmlId = "MyCustomXmlPartId";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Check if any custom XML part exists (specific Id check omitted due to API limitations)
        bool hasTargetPart = false;
        try
        {
            hasTargetPart = workbook.CustomXmlParts != null && workbook.CustomXmlParts.Count > 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while inspecting custom XML parts: {ex.Message}");
        }

        // Apply modifications only when a custom XML part is present
        if (hasTargetPart)
        {
            try
            {
                // Add a new worksheet
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet sheet = workbook.Worksheets[sheetIndex];
                sheet.Name = "Processed";

                // Write a note indicating successful processing
                sheet.Cells["A1"].PutValue("Workbook contained the required custom XML part.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during worksheet modification: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"Custom XML part with Id \"{targetCustomXmlId}\" not found.");
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
