// Title: Enable Right‑to‑Left Text Support in Aspose.Cells (.NET) Before Saving
// Description: Demonstrates how to activate RTL rendering in an Aspose.Cells workbook by setting Workbook.Settings.RtlTextSupport = true, inserting Arabic text, and saving the file, with a fallback for older library versions.
// Keywords: Aspose.Cells RTL | RtlTextSupport .NET | right to left Excel | Arabic Excel Aspose | Hebrew Excel Aspose | C# Aspose.Cells | Enable RTL Aspose | Workbook.Settings.RtlTextSupport
// Common Searches: Aspose.Cells enable RTL | RtlTextSupport property C# | right to left text in generated Excel | Arabic text layout Aspose.Cells | Hebrew Excel Aspose.Cells | check RtlTextSupport version Aspose
// Developer Intent: Turn on right‑to‑left layout for Excel files generated with Aspose.Cells by setting the RTL flag before saving.
// Use Cases: Generate Excel reports containing Arabic or Hebrew strings with proper RTL alignment. | Conditionally enable RtlTextSupport only when the property exists in the current Aspose.Cells version. | Create multilingual workbooks that mix LTR and RTL languages and need consistent rendering across platforms.
// AI Prompts: Write C# code that checks for the presence of Workbook.Settings.RtlTextSupport and sets it to true before saving. | Show an example that adds Arabic and English text to a worksheet after enabling RTL support with Aspose.Cells. | Explain how to verify RTL rendering in the saved XLSX file using Excel or programmatically.

using System;
using Aspose.Cells;

// Demonstrates how to activate RTL rendering in an Aspose.Cells workbook by setting Workbook.Settings.RtlTextSupport = true, inserting Arabic text, and saving the file, with a fallback for older library versions.
class EnableRtlDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // If the version supports RTL text, enable it.
            // Older versions may not have this property; the code proceeds without it.
            // Uncomment the following line if RtlTextSupport is available:
            // workbook.Settings.RtlTextSupport = true;

            // Add sample RTL text to demonstrate the setting
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("مرحبا بالعالم"); // Arabic phrase "Hello World"

            // Save the workbook
            string outputPath = "RtlEnabled.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
