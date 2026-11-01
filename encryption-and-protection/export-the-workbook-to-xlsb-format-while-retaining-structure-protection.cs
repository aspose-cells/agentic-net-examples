// Title: Export an XLSX workbook to XLSB while preserving structure protection with Aspose.Cells for .NET
// Common Searches: how to convert an .xlsx file to .xlsb and keep workbook structure protection using Aspose.Cells C# | Aspose.Cells C# protect workbook structure before saving as binary XLSB | save workbook as XLSB with password‑protected structure in .NET | C# Aspose.Cells example for protecting workbook structure and exporting to XLSB
// Tags: Aspose.Cells protect workbook structure C# | export workbook to XLSB Aspose.Cells | Workbook.Protect with ProtectionType.Structure | SaveFormat.Xlsb binary Excel export | convert XLSX to XLSB preserving protection

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, applies structure protection with Workbook.Protect using a password, and saves the workbook as a binary XLSB file via SaveFormat.Xlsb, including basic exception handling.
class ExportWorkbook
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsb";
            const string password = "MyPassword";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, password);

            // Save the workbook in XLSB format while retaining the structure protection
            workbook.Save(outputPath, SaveFormat.Xlsb);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
