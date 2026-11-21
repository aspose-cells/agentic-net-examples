// Title: Add a registered Scripting library reference to a VBA project in a macro‑enabled workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add registered reference to VBA project in .xlsm workbook | How to include Scripting library in Excel macro with Aspose.Cells VbaProject | C# example for VbaProject.References.AddRegisteredReference method | Saving macro‑enabled workbook after adding VBA library reference using Aspose.Cells | Programmatically create VBA project and add library reference in .NET
// Tags: Aspose.Cells VbaProject add registered reference | C# add VBA library reference Aspose.Cells | macro-enabled workbook save with VBA references | Add Scripting library to Excel VBA project | VbaProject References.AddRegisteredReference example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject

// The sample creates or loads a workbook, accesses its VbaProject, adds a registered "Scripting" library reference using VbaProject.References.AddRegisteredReference, ensures the output directory exists, and saves the file as a macro‑enabled .xlsm workbook while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project (Aspose.Cells creates one automatically if missing)
            VbaProject vbaProject = workbook.VbaProject;

            // Add a registered library reference to the VBA project.
            // The second argument is the library ID (LibID). For many built‑in libraries an empty string works.
            vbaProject.References.AddRegisteredReference("Scripting", "");

            // Define output file path
            string outputPath = "OutputWithReference.xlsm";

            // Ensure the directory for the output file exists (handle cases where the path has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in a macro‑enabled format to preserve the VBA project.
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
