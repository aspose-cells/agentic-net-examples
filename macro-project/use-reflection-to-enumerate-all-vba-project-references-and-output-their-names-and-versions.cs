// Title: C# – Enumerate VBA project references and display their names and versions using Aspose.Cells and reflection
// Common Searches: C# Aspose.Cells enumerate VBA references in an .xlsm workbook | how to retrieve VBA reference name and version with reflection in .NET | list VBA project library references using Aspose.Cells C# code | read VBA References collection from Excel macro file programmatically | using reflection to access VBA project references in Aspose.Cells
// Tags: reflect VBA References property Aspose.Cells | enumerate VBA project libraries .NET | output VBA reference name version C# | access VBA project references via reflection | Aspose.Cells VBA project inspection

using System;
using System.Collections;
using System.Reflection;
using Aspose.Cells;

// The example loads an .xlsm workbook, checks for an embedded VBA project, uses .NET reflection to obtain the public `References` collection, iterates through each reference object, reads its `Name` and `Version` properties via reflection, and writes the name‑version pairs to the console.
class Program
{
    static void Main()
    {
        // Load the workbook that contains a VBA project.
        // Adjust the file path as needed.
        var workbook = new Workbook("input.xlsm");

        // Get the VBA project from the workbook.
        var vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("No VBA project found in the workbook.");
            return;
        }

        // Use reflection to obtain the 'References' property of the VBA project.
        Type vbaProjectType = vbaProject.GetType();
        PropertyInfo referencesProp = vbaProjectType.GetProperty(
            "References",
            BindingFlags.Public | BindingFlags.Instance);

        if (referencesProp == null)
        {
            Console.WriteLine("Unable to locate the References property via reflection.");
            return;
        }

        // Retrieve the collection of references.
        object referencesObj = referencesProp.GetValue(vbaProject);
        if (referencesObj is not IEnumerable references)
        {
            Console.WriteLine("References property is not enumerable.");
            return;
        }

        // Enumerate each reference and output its Name and Version using reflection.
        foreach (object reference in references)
        {
            Type refType = reference.GetType();

            PropertyInfo nameProp = refType.GetProperty(
                "Name",
                BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo versionProp = refType.GetProperty(
                "Version",
                BindingFlags.Public | BindingFlags.Instance);

            string name = nameProp?.GetValue(reference) as string ?? "<Unnamed>";
            string version = versionProp?.GetValue(reference) as string ?? "<NoVersion>";

            Console.WriteLine($"{name} - {version}");
        }
    }
}
