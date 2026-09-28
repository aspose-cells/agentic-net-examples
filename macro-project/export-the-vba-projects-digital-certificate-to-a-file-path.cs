// Title: Export a VBA project's digital certificate to a .cer file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, checks for a VBA project, and saves its digital signature as a .cer file with Aspose.Cells. | Show how to use reflection in C# to obtain the DigitalSignature object of a VBA project and call its Export method via Aspose.Cells. | Write a C# snippet that gracefully handles missing files, absent VBA projects, and missing digital signatures while exporting the certificate to a given path.
// Common Searches: aspnet extract vba macro certificate from excel using aspose.cells | c# get vba project digital signature and save as .cer file | how to use reflection to export VBA digital signature with Aspose.Cells | Aspose.Cells example for exporting VBA certificate in C# | handle missing vba project when exporting certificate c#
// Tags: Aspose.Cells VBA certificate export | C# reflection DigitalSignature API | save VBA certificate as .cer | load Excel workbook with VBA project | detect absent VBA signature C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, verifies that it contains a VBA project, uses reflection to access the project's DigitalSignature object, and invokes its Export method to write the certificate to a .cer file. It includes error handling for missing files, absent VBA projects, and missing digital signatures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string certPath = "VbaCertificate.cer";

        try
        {
            // Ensure the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that may contain a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Check for a VBA project
            var vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Use reflection to access the DigitalSignature property (available in newer versions)
            var digitalSignatureProp = vbaProject.GetType().GetProperty("DigitalSignature");
            if (digitalSignatureProp == null)
            {
                Console.WriteLine("The VBA project does not contain a digital signature (property not found).");
                return;
            }

            var digitalSignature = digitalSignatureProp.GetValue(vbaProject);
            if (digitalSignature == null)
            {
                Console.WriteLine("The VBA project does not contain a digital signature.");
                return;
            }

            // Use reflection to call the Export method of the DigitalSignature object
            var exportMethod = digitalSignature.GetType().GetMethod("Export");
            if (exportMethod == null)
            {
                Console.WriteLine("Export method not found on DigitalSignature object.");
                return;
            }

            exportMethod.Invoke(digitalSignature, new object[] { certPath });
            Console.WriteLine("Digital certificate exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
