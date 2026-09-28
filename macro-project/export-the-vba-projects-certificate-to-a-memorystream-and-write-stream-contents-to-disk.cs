// Title: Export a signed VBA project's certificate to a binary file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsm workbook with Aspose.Cells, verifies the VBA project is signed, extracts the certificate bytes via reflection, and saves them to a .bin file. | Show how to access the VbaProject.Signature property in Aspose.Cells, pipe the certificate data through a MemoryStream, and write the result to disk.
// Common Searches: aspocells export signed vba certificate to file | c# retrieve VBA project signature bytes from xlsm workbook | how to save VBA macro certificate as binary using Aspose.Cells | extract signed macro certificate with reflection in .NET | write VBA project certificate to disk from Aspose.Cells workbook
// Tags: export VBA certificate Aspose.Cells | retrieve VbaProject signature bytes | write certificate binary C# | validate signed macro before extraction | copy certificate via MemoryStream

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject

// The example loads an .xlsm workbook, checks that it contains a signed VBA project, uses reflection to obtain the certificate bytes from the VbaProject.Signature property, streams the data through a MemoryStream, and writes the bytes to a binary file (VbaCertificate.bin).
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "VbaCertificate.bin";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the VBA project
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project (may be null if no VBA project exists)
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Check if the VBA project is signed
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed; no certificate to export.");
                return;
            }

            // Attempt to retrieve the certificate bytes via reflection (compatible with multiple versions)
            byte[] certificateBytes = null;
            PropertyInfo signatureProp = typeof(VbaProject).GetProperty("Signature", BindingFlags.Public | BindingFlags.Instance);
            if (signatureProp != null)
            {
                certificateBytes = signatureProp.GetValue(vbaProject) as byte[];
            }

            if (certificateBytes == null || certificateBytes.Length == 0)
            {
                Console.WriteLine("Unable to retrieve the VBA project certificate.");
                return;
            }

            // Export the certificate to a file
            using (MemoryStream memoryStream = new MemoryStream(certificateBytes))
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                memoryStream.CopyTo(fileStream);
            }

            Console.WriteLine("VBA project certificate exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
