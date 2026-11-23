// Title: Copy VBA UserForms from a source .xlsm template to a new workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy VBA UserForms from one .xlsm to another | How to preserve VBA macros when saving a workbook with Aspose.Cells | Duplicate Excel template with its UserForms using Aspose.Cells in .NET | Save a workbook as macro‑enabled XLSM while keeping UserForms via Aspose.Cells | Load a macro‑enabled workbook and retain the VBA project with Aspose.Cells C#
// Tags: copy userforms aspose.cells c# | preserve vba project aspose.cells | save workbook as xlsm aspose.cells | duplicate macro-enabled workbook aspose.cells | load vba enabled workbook aspose.cells | transfer excel userforms .net

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a source macro‑enabled workbook, loads it with Aspose.Cells, and saves it to a new .xlsm file, ensuring the VBA project and all embedded UserForms are retained.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceTemplate.xlsm";
            const string destPath = "DestinationWithUserForms.xlsm";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook that contains the UserForms
            Workbook sourceWb = new Workbook(sourcePath);

            // Save the workbook to the destination path (preserves VBA project and UserForms)
            sourceWb.Save(destPath, SaveFormat.Xlsm);
            Console.WriteLine($"Destination workbook saved: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
