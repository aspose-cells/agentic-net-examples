// Title: How to add a registered Scripting library reference to a VBA project in a macro-enabled Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add registered reference to VBA project in .xlsm file | C# code to include Scripting library in Excel macro workbook with Aspose.Cells | Using Aspose.Cells VbaProject.References.AddRegisteredReference to register a library | Programmatically add a VBA library reference in a macro-enabled workbook with Aspose.Cells | Add Scripting reference to VBA project via Aspose.Cells .NET API
// Tags: Aspose.Cells VbaProject AddRegisteredReference | register VBA library reference C# | macro-enabled workbook library registration | include scripting library in VBA project | C# Aspose.Cells VBA references API

using System;
using Aspose.Cells;
using System.IO;

// The example creates a macro-enabled Workbook, accesses its VbaProject, registers the Scripting library using References.AddRegisteredReference, and saves the file as an .xlsm workbook, with exception handling for any errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that can contain macros
            var workbook = new Workbook();

            // Access the VBA project associated with the workbook
            var vbaProject = workbook.VbaProject;

            // Add a registered library reference (e.g., "Scripting") to the VBA project
            // Both name and libid are required by the API
            vbaProject.References.AddRegisteredReference("Scripting", "Scripting");

            // Define the output file path
            string outputPath = "output.xlsm";

            // Save the workbook with macro support
            workbook.Save(outputPath, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
