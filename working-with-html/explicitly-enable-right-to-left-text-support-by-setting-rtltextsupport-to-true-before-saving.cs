// Title: C# – Enable Right‑to‑Left (RTL) Text Support in Aspose.Cells Workbook Before Save
// Description: Creates a new Workbook, detects the RTL property (IsRightToLeft or RightToLeft) via reflection for version‑agnostic support, sets it to true, writes Arabic sample text to A1, guarantees the output folder exists, and saves the file as XLSX.
// Keywords: Aspose.Cells RTL support C# | right‑to‑left text Excel Aspose | IsRightToLeft property | RightToLeft setting Aspose.Cells | Arabic Excel generation .NET | Hebrew Excel Aspose.Cells
// Common Searches: how to turn on RTL in Aspose.Cells .NET | set right‑to‑left direction for Excel workbook using Aspose | reflection for IsRightToLeft vs RightToLeft Aspose.Cells | enable Arabic text in generated Excel with Aspose
// Developer Intent: Activate right‑to‑left text rendering in a workbook before it is saved.
// Use Cases: Produce Arabic or Hebrew financial reports where cell direction must be RTL. | Create multilingual invoices that mix left‑to‑right and right‑to‑left sheets. | Export workbooks to HTML or PDF while preserving RTL layout for specific languages.
// AI Prompts: Generate C# code that enables RTL text in Aspose.Cells, handling both IsRightToLeft and RightToLeft properties via reflection. | Show how to verify RTL rendering after saving a workbook to XLSX, HTML, and PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Workbook, detects the RTL property (IsRightToLeft or RightToLeft) via reflection for version‑agnostic support, sets it to true, writes Arabic sample text to A1, guarantees the output folder exists, and saves the file as XLSX.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable right-to-left text support (use reflection to handle version differences)
            var settings = workbook.Settings;
            var rtlProp = settings.GetType().GetProperty("IsRightToLeft") ??
                          settings.GetType().GetProperty("RightToLeft");

            if (rtlProp != null && rtlProp.CanWrite)
            {
                rtlProp.SetValue(settings, true);
            }

            // Add sample right-to-left text (Arabic) to demonstrate the setting
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("مثال نص من اليمين إلى اليسار");

            // Define output file path
            string outputPath = "RtlSupportDemo.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
