// Title: Convert an XLSX workbook to XLSB while preserving structure protection with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export XLSX to XLSB keeping workbook structure protection | How to save a protected Excel workbook as binary (XLSB) using Aspose.Cells .NET | Aspose.Cells protect workbook structure then convert to XLSB in C# | Preserve workbook protection when converting XLSX to XLSB with Aspose.Cells | Save Excel file as XLSB with password‑protected structure using Aspose.Cells
// Tags: Aspose.Cells workbook.Protect API | XLSB file generation with Aspose.Cells | C# Excel binary format conversion | Maintain protection on XLSB output | Apply structure lock during XLSB save

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, applies structure protection with a password via workbook.Protect, and then saves the workbook as an XLSB file using Aspose.Cells, ensuring the protection remains intact while handling file existence checks and exceptions.
class ExportToXlsb
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsb";
        const string password = "myPassword";

        try
        {
            // Ensure the source file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, password);

            // Save the workbook in XLSB format, preserving protection
            workbook.Save(outputPath, SaveFormat.Xlsb);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
