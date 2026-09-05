// Title: Create an XLSX workbook in C# with Aspose.Cells, set a strong opening password, and save with built-in encryption
// AI Prompts: Write C# code using Aspose.Cells to generate a new workbook, assign a strong opening password, and save it as an encrypted XLSX file. | Show how to change the workbook password after creation and re‑save the file while preserving encryption with Aspose.Cells.
// Common Searches: asp.net set opening password for excel file using Aspose.Cells | c# protect xlsx workbook with password and encryption Aspose.Cells | how to save encrypted Excel workbook with Aspose.Cells .NET | default encryption strength of password‑protected XLSX in Aspose.Cells | create workbook and ensure output folder exists before saving with Aspose.Cells
// Tags: Aspose.Cells password protection for XLSX | C# encrypt Excel workbook using Aspose.Cells | Workbook.Settings.Password usage | SaveFormat.Xlsx with encryption | ensure output directory exists Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, assigns a strong opening password via Workbook.Settings.Password, ensures the target directory exists, and saves the file as an XLSX workbook. Aspose.Cells automatically applies strong encryption to the password‑protected file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set a password to open the workbook (default encryption will be applied)
            workbook.Settings.Password = "StrongPassword!2026";

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
