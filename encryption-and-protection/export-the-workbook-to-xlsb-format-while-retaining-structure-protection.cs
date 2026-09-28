// Title: Export an Excel workbook to XLSB while preserving structure protection using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file, applies workbook structure protection with a password, and saves it as an .xlsb file using Aspose.Cells. | Show how to protect the workbook structure before converting to XLSB format with Aspose.Cells in a .NET application. | Provide a step‑by‑step example for converting an existing workbook to XLSB while retaining its structure password using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells export workbook to XLSB preserving structure password | How to keep workbook structure protection when saving as XLSB in .NET | Convert .xlsx to .xlsb with Aspose.Cells and retain protection | Aspose.Cells protect workbook structure before saving as binary Excel
// Tags: Aspose.Cells protect structure API | XLSB format export with Aspose.Cells | structure password retention during XLSB save | C# workbook protection before binary export | Aspose.Cells workbook conversion preserving security

using System;
using System.IO;
using Aspose.Cells;

// The example loads source.xlsx, applies structure protection with a password, and saves the workbook as output.xlsb using Aspose.Cells for .NET, including checks for missing files and error handling.
class ExportToXlsb
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string outputPath = "output.xlsb";
        const string password = "myPassword";

        // Verify source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(sourcePath);

            // Protect workbook structure with a password (retained when saved as XLSB)
            workbook.Protect(ProtectionType.Structure, password);

            // Save the workbook in XLSB format
            workbook.Save(outputPath, SaveFormat.Xlsb);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving, protection)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
