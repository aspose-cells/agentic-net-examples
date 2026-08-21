// Title: Enumerate VBA Project References and Versions with Reflection – Aspose.Cells for .NET
// Description: Loads a macro‑enabled workbook, checks for a VBA project, iterates its References collection, and uses .NET reflection to read each VbaProjectReference's Name and Libid. The Libid string is parsed to extract the version component, which is printed alongside the reference name. The workbook is then saved unchanged.
// Keywords: Aspose.Cells | VBA reference enumeration | C# reflection | Libid version extraction | macro-enabled workbook | list VBA references | Excel COM libraries | VbaProjectReference | .NET Excel automation
// Common Searches: list VBA references Aspose.Cells C# | extract version from VBA Libid string | use reflection to read VBA reference properties | how to enumerate VBA project references in .xlsm | retrieve VBA library GUID with Aspose.Cells
// Developer Intent: Display every VBA reference name and its version from a macro‑enabled workbook using reflection.
// Use Cases: Detect missing or outdated COM libraries before deploying Excel macros. | Create an inventory of external VBA libraries across multiple workbooks for compliance audits. | Log reference details during batch processing of .xlsm files.
// AI Prompts: Generate C# code that loads an .xlsm file with Aspose.Cells and prints each VBA reference's Name and version extracted from Libid using reflection. | Write a robust method to parse the Libid string and safely return the version segment, handling unexpected formats. | Show how to extend the sample to also output the GUID of each VBA reference via reflection.

using System;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a macro‑enabled workbook, checks for a VBA project, iterates its References collection, and uses .NET reflection to read each VbaProjectReference's Name and Libid. The Libid string is parsed to extract the version component, which is printed alongside the reference name. The workbook is then saved unchanged.
class EnumerateVbaReferences
{
    static void Main()
    {
        // Load a macro-enabled workbook that contains a VBA project
        string inputPath = "input.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Check if the workbook has a VBA project
        if (workbook.VbaProject != null)
        {
            // Get the collection of VBA references
            VbaProjectReferenceCollection references = workbook.VbaProject.References;
            Console.WriteLine($"Total VBA references: {references.Count}");

            // Use reflection to obtain the Name and Libid properties
            Type refType = typeof(VbaProjectReference);
            PropertyInfo nameProp = refType.GetProperty("Name");
            PropertyInfo libidProp = refType.GetProperty("Libid");

            // Enumerate each reference and output its name and version (extracted from Libid)
            for (int i = 0; i < references.Count; i++)
            {
                object refObj = references[i];
                string name = nameProp?.GetValue(refObj) as string;
                string libid = libidProp?.GetValue(refObj) as string;

                // Attempt to extract a version component from the Libid string
                string version = "N/A";
                if (!string.IsNullOrEmpty(libid))
                {
                    string[] parts = libid.Split('#');
                    if (parts.Length > 1)
                    {
                        version = parts[1];
                    }
                }

                Console.WriteLine($"Reference {i + 1}: Name = {name}, Version = {version}");
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }

        // Save the workbook (unchanged) as a macro-enabled file
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
