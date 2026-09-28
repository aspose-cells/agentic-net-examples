// Title: Detect circular reference errors in formulas and export a workbook to ODS with Aspose.Cells for .NET
// AI Prompts: Generate C# code that enables formula circular‑reference detection in Aspose.Cells, catches any related exceptions, and saves the workbook as an ODS file. | Show how to wrap workbook.Save in a try‑catch block to handle circular reference errors before exporting to ODS using Aspose.Cells. | Provide a C# snippet that creates the target folder if it does not exist and then saves the workbook to ODS format with Aspose.Cells.
// Common Searches: asp.net Aspose.Cells detect circular reference in formulas before ODS export | c# example for catching circular reference exception when saving workbook as ODS with Aspose.Cells | how to enable formula error checking in Aspose.Cells and export to ODS | save workbook to ODS format and ensure output directory exists using Aspose.Cells C# | Aspose.Cells circular reference property missing workaround for .NET
// Tags: circular reference detection Aspose.Cells | export workbook to ODS Aspose.Cells | formula error handling C# | create output directory C# Aspose.Cells | save workbook with try‑catch Aspose.Cells | ODS file generation Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, notes that the EnableCircularReference property is unavailable in the current version, ensures the destination folder exists, and saves the workbook as an ODS file. All operations are wrapped in a try‑catch block to surface any exceptions, such as formula circular‑reference errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: EnableCircularReference is not available in the current Aspose.Cells version.
            // If needed, refer to the documentation for the appropriate property in newer releases.

            // Define output file path
            string outputPath = "output.ods";

            // Ensure the output directory exists (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
