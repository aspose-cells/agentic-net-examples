// Title: Copy worksheets from a regular .xlsx workbook into a macro‑enabled .xlsm workbook while preserving the existing VBA project using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens a macro‑enabled .xlsm file and a standard .xlsx file, copies every worksheet from the .xlsx into the .xlsm, and saves the result as a new .xlsm keeping all VBA macros intact. | Show how to merge a non‑macro workbook into a macro‑enabled workbook without breaking the VBA project, using Aspose.Cells in C#. | Provide a C# snippet that checks file existence, loads both workbooks, iterates through the destination workbook’s worksheets, adds them to the source macro‑enabled workbook, and saves the combined file as Xlsm while preserving macro security settings.
// Common Searches: asp.net copy worksheets into macro enabled workbook preserving VBA Aspose.Cells | how to keep VBA project when merging .xlsx and .xlsm files with Aspose.Cells | C# Aspose.Cells merge workbooks without losing macros | save combined Excel file as macro enabled Xlsm using Aspose.Cells | load macro enabled workbook and add sheets from regular workbook C#
// Tags: copy worksheets Aspose.Cells | preserve VBA project Aspose.Cells | merge macro enabled workbook C# | save as Xlsm Aspose.Cells | load macro enabled workbook C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads a macro‑enabled source workbook and a regular destination workbook, copies all worksheets from the destination into the source (retaining the original VBA project), and saves the merged workbook as a new .xlsm file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsm";
            const string destinationPath = "destination.xlsx";
            const string outputPath = "destination_with_macros.xlsm";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source macro‑enabled workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Load the destination workbook if it exists; otherwise create a new workbook
            Workbook destinationWorkbook = File.Exists(destinationPath)
                ? new Workbook(destinationPath)
                : new Workbook();

            // Copy all worksheets from the destination workbook into the source workbook
            // This preserves the VBA project already present in the source workbook
            foreach (Worksheet ws in destinationWorkbook.Worksheets)
            {
                // Add a copy of the worksheet to the source workbook
                sourceWorkbook.Worksheets.AddCopy(ws.Index);
            }

            // Save the combined workbook as a macro‑enabled file
            sourceWorkbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
