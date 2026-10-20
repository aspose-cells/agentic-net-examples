// Title: Use reflection in C# with Aspose.Cells to enumerate VBA project references and show each reference’s name and version in an .xlsm workbook
// Common Searches: c# how to list VBA references and their versions from an .xlsm file using Aspose.Cells | using .net reflection to read VBA project reference names in Excel workbook | retrieve VBA reference major and minor version numbers with Aspose.Cells C# example | enumerate VBA project references programmatically in a macro-enabled workbook | Aspose.Cells C# code to display VBA reference details via reflection
// Tags: csharp reflection enumerate vba references | aspose.cells vba project reference version extraction | xlsm workbook vba reference enumeration | list vba references using .net reflection | retrieve vba reference name and version c#

using System;
using System.Reflection;
using Aspose.Cells;

// The sample loads an .xlsm workbook with Aspose.Cells, accesses its VbaProject, uses .NET reflection to obtain the public References collection, iterates through each reference, and prints the reference name together with its major and minor version numbers.
class Program
{
    static void Main()
    {
        // Load the workbook that contains a VBA project (replace with your file path)
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project
        var vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("No VBA project found in the workbook.");
            return;
        }

        // Use reflection to obtain the 'References' property of the VBA project
        PropertyInfo referencesProp = vbaProject.GetType()
                                                .GetProperty("References", BindingFlags.Public | BindingFlags.Instance);
        if (referencesProp == null)
        {
            Console.WriteLine("Unable to locate the References property via reflection.");
            return;
        }

        // Retrieve the collection of references
        object referencesObj = referencesProp.GetValue(vbaProject, null);
        if (referencesObj == null)
        {
            Console.WriteLine("The VBA project contains no references.");
            return;
        }

        // Enumerate the references using reflection
        var enumerable = referencesObj as System.Collections.IEnumerable;
        foreach (object reference in enumerable)
        {
            Type refType = reference.GetType();

            // Get the reference name
            PropertyInfo nameProp = refType.GetProperty("Name");
            string name = nameProp?.GetValue(reference, null)?.ToString() ?? "Unnamed";

            // Get the major and minor version numbers
            PropertyInfo majorProp = refType.GetProperty("Major");
            PropertyInfo minorProp = refType.GetProperty("Minor");
            int major = majorProp != null ? (int)majorProp.GetValue(reference, null) : 0;
            int minor = minorProp != null ? (int)minorProp.GetValue(reference, null) : 0;

            Console.WriteLine($"Reference: {name}, Version: {major}.{minor}");
        }
    }
}
