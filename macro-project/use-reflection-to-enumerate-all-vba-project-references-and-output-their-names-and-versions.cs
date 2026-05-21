using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Reflection;

class EnumerateVbaReferences
{
    static void Main()
    {
        // Load a macro-enabled workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsm");

        // Get the VBA project via reflection
        PropertyInfo vbaProjectProp = typeof(Workbook).GetProperty("VbaProject");
        VbaProject vbaProject = (VbaProject)vbaProjectProp.GetValue(workbook);

        if (vbaProject == null)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Obtain the References collection using reflection
        PropertyInfo referencesProp = typeof(VbaProject).GetProperty("References");
        VbaProjectReferenceCollection references = (VbaProjectReferenceCollection)referencesProp.GetValue(vbaProject);

        // Prepare reflection objects for the properties we need
        PropertyInfo nameProp = typeof(VbaProjectReference).GetProperty("Name");
        PropertyInfo libidProp = typeof(VbaProjectReference).GetProperty("Libid");

        // Enumerate all references and output their name and version (extracted from Libid)
        for (int i = 0; i < references.Count; i++)
        {
            VbaProjectReference reference = references[i];
            string name = (string)nameProp.GetValue(reference);
            string libid = (string)libidProp.GetValue(reference);

            // Attempt to extract version information from the Libid string (if present)
            string version = "Unknown";
            if (!string.IsNullOrEmpty(libid))
            {
                // Libid format often contains version after the second '#'
                string[] parts = libid.Split('#');
                if (parts.Length > 2)
                {
                    version = parts[2];
                }
            }

            Console.WriteLine($"Reference {i + 1}: Name = {name}, Version = {version}");
        }
    }
}