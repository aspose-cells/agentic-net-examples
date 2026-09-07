// Title: Enumerate VBA project references in an .xlsm workbook and display their names and version numbers using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a macro‑enabled Excel file with Aspose.Cells, accesses the workbook’s VbaProject, and prints each VBA reference’s Name together with its MajorVersion and MinorVersion using reflection or dynamic typing. | Extend the sample to also retrieve and output the GUID and FullPath properties of each VBA reference in the workbook.
// Common Searches: aspocells c# list vba references and their versions in xlsm | how to get VBA reference names from a macro-enabled workbook using Aspose.Cells | C# reflection to read VBA project reference version numbers with Aspose.Cells | retrieve VBA reference GUID from Excel file programmatically in .NET
// Tags: enumerate vba references Aspose.Cells | read vba reference version .NET | list macro project references C# | access vba project references using reflection | extract vba reference guid Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example verifies an .xlsm file exists, loads it with Aspose.Cells, obtains its VbaProject, iterates over the References collection via dynamic reflection, and prints each reference’s Name and version (MajorVersion.MinorVersion), handling missing data and errors.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsm";

        // Verify that the input file exists.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook that contains a VBA project.
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Retrieve the VBA project from the workbook.
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("No VBA project found in the workbook.");
            return;
        }

        // Get the collection of VBA references (use non‑specific types to avoid missing assemblies).
        var references = vbaProject.References;
        if (references == null || ((ICollection)references).Count == 0)
        {
            Console.WriteLine("No VBA references found.");
            return;
        }

        // Enumerate each reference and output its name and version.
        foreach (object refObj in (IEnumerable)references)
        {
            try
            {
                // Use dynamic to access members without needing the concrete type at compile time.
                dynamic reference = refObj;
                string name = reference.Name ?? "Unnamed";
                int major = reference.MajorVersion;
                int minor = reference.MinorVersion;
                Console.WriteLine($"{name} - Version {major}.{minor}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing a reference: {ex.Message}");
            }
        }
    }
}
