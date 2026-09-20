// Title: Copy an XLSM workbook while preserving VBA macros and UserForm storage using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a macro‑enabled XLSM workbook, sets Settings.EnableMacros, and saves it as a new XLSM file so that all VBA modules and UserForms remain intact. | Show how to duplicate a macro‑enabled Excel file with Aspose.Cells without stripping any VBA code or UserForm data.
// Common Searches: how to keep VBA UserForms when duplicating an xlsm file with Aspose.Cells | Aspose.Cells .NET copy macro-enabled workbook without losing VBA code | C# save workbook as xlsm while preserving macros and forms | clone an xlsm workbook preserving all VBA components using Aspose.Cells
// Tags: copy macro-enabled XLSM using Aspose.Cells | retain VBA UserForms Aspose.Cells .NET | enable macro support SaveFormat.Xlsm | clone workbook with VBA modules intact | macro-enabled workbook cloning .NET

using Aspose.Cells;
using System;
using System.IO;

// // Loads a macro‑enabled XLSM workbook, enables macro support, and saves it as a new XLSM file, preserving all VBA macros and UserForm components.
class PreserveMacrosAndUserForms
{
    static void Main()
    {
        try
        {
            string sourcePath = "SourceWithMacros.xlsm";
            string destPath = "DestinationPreserved.xlsm";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook; macros and UserForms are retained automatically
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Enable macro support when saving
            sourceWorkbook.Settings.EnableMacros = true;

            // Save the workbook as XLSM, preserving all VBA components
            sourceWorkbook.Save(destPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
