// Title: Add a registered library reference to a VBA project in a macro‑enabled workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsm file with Aspose.Cells, retrieves its VbaProject, and calls References.AddRegisteredReference to add a COM library such as stdole, then saves the workbook preserving macros. | Show how to programmatically insert a registered VBA library reference into a macro‑enabled Excel workbook using the Aspose.Cells VbaProject API in C#. | Demonstrate updating VBA references in a loaded workbook (e.g., adding stdole) with Aspose.Cells and exporting the file as Xlsm.
// Common Searches: asp.net add stdole reference to VBA project in existing xlsm using Aspose.Cells | C# example for References.AddRegisteredReference in Aspose.Cells VbaProject | how to modify VBA references in a macro‑enabled workbook with Aspose.Cells | Aspose.Cells add registered COM library to VBA project in .xlsm file | preserve macros after adding VBA library reference using Aspose.Cells C#
// Tags: add registered VBA library reference Aspose.Cells | VbaProject References.AddRegisteredReference C# | modify VBA references in macro-enabled workbook | Aspose.Cells add stdole COM reference | load and save .xlsm with VBA project using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads an existing .xlsm workbook, accesses its VbaProject, adds a registered library reference (e.g., stdole) via References.AddRegisteredReference, and saves the file as a macro‑enabled workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsm";
            const string outputPath = "output.xlsm";

            // Ensure the template file exists; it must contain a VBA project.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Required template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook that already has a VBA project.
            Workbook workbook = new Workbook(templatePath);

            // Access the existing VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Add a registered library reference (both name and libid are required).
            vbaProject.References.AddRegisteredReference("stdole", "stdole");

            // Save as macro‑enabled workbook to retain VBA.
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
