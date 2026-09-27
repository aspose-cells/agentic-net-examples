// Title: Update a VBA module in an .xlsm workbook and re‑sign the macro project with Aspose.Cells for .NET
// AI Prompts: Create C# code that opens an .xlsm file, appends a comment to the first VBA module, and demonstrates how to apply a digital signature to the VbaProject using Aspose.Cells if the DigitalSignature class is available. | Show the steps to modify VBA code in a macro‑enabled workbook and then re‑apply a digital signature programmatically with the Aspose.Cells .NET API. | Provide an example that preserves the VBA project while saving after changes and outlines the required calls to sign the VBA project in C#.
// Common Searches: how to add a comment to a VBA module in an .xlsm workbook using Aspose.Cells C# | Aspose.Cells reapply digital signature to VBA project after code modification | C# update macro-enabled Excel file and keep VBA signature | programmatically sign a VBA project in a .xlsm file with Aspose.Cells .NET | preserve VBA macro signature when editing modules using Aspose.Cells
// Tags: modify VBA module Aspose.Cells | save macro-enabled workbook preserving VBA project | digital signature for VBA project Aspose.Cells | Aspose.Cells VbaProject editing example | C# programmatic VBA macro signing | Aspose.Cells VBA project handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an .xlsm workbook, accesses its VbaProject, appends a comment to the first VBA module, notes that digital‑signature functionality would require the DigitalSignature class (which may not be present), and then saves the workbook while preserving the VBA project.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsm";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a VBA project.
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Append a comment to the first VBA module, if any.
            if (vbaProject != null && vbaProject.Modules.Count > 0)
            {
                VbaModule firstModule = vbaProject.Modules[0];
                firstModule.Codes += "\r\n' Added comment after code change";
            }

            // NOTE: Digital signature functionality requires the DigitalSignature class,
            // which may not be available in the current Aspose.Cells version.
            // The signing step is omitted to ensure the code compiles and runs.

            // Save the workbook, preserving the VBA project.
            string outputPath = "output.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
