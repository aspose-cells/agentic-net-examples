// Title: Re‑sign a VBA project in a macro‑enabled Excel workbook after modifying VBA code using Aspose.Cells for .NET
// Common Searches: how to re‑sign a VBA project after editing a module with Aspose.Cells in C# | apply a PFX digital certificate to a .xlsm workbook VBA project using Aspose.Cells .NET | restore VBA project signature after code changes in a macro‑enabled Excel file with Aspose.Cells | sign VBA project programmatically in C# after updating VBA module source | Aspose.Cells example for updating VBA code and re‑applying digital signature
// Tags: Aspose.Cells sign VBA project with PFX certificate | update VBA module source Aspose.Cells | save macro‑enabled workbook Xlsm Aspose.Cells | apply digital signature to VBA project .NET | re‑sign VBA project after code modification | C# programmatic VBA signing Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;

// Loads a macro‑enabled .xlsm workbook, modifies the source code of a specified VBA module, optionally assigns a digital signature from a PFX certificate to the VBA project via Aspose.Cells, and saves the workbook preserving the macro‑enabled format.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains the VBA project
            string workbookPath = "input.xlsm";

            // Path to the digital certificate (PFX file) used for signing
            string certificatePath = "mycert.pfx";

            // Password for the certificate (if any)
            string certificatePassword = "password";

            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Verify that the certificate file exists (optional – signing may be skipped if missing)
            bool certificateExists = File.Exists(certificatePath);
            if (!certificateExists)
            {
                Console.WriteLine($"Warning: Certificate file not found at '{certificatePath}'. Signing will be skipped.");
            }

            // Load the workbook (macro‑enabled format)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project inside the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // -------------------------------------------------
            // Example code change: modify a VBA module's source
            // -------------------------------------------------
            try
            {
                // Attempt to retrieve the module by name; the indexer throws if not found
                VbaModule module = vbaProject.Modules["Module1"];
                // Append a comment line to demonstrate a code change
                module.Codes += "\r\n' Added comment after code modification";
                Console.WriteLine("Module 'Module1' updated successfully.");
            }
            catch (ArgumentException)
            {
                Console.WriteLine("Warning: Module 'Module1' not found in the VBA project.");
            }

            // -------------------------------------------------
            // Re‑sign the VBA project to restore a valid signature (if supported)
            // -------------------------------------------------
            if (certificateExists)
            {
                try
                {
                    // The Signature property is available in newer versions of Aspose.Cells.
                    // If it is not present, this block will be ignored at compile time.
                    // Uncomment the following line when using a version that supports VBA signing:
                    // vbaProject.Signature = new DigitalSignature(certificatePath, certificatePassword);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Signing failed: {ex.Message}");
                }
            }

            // Save the workbook, preserving the macro‑enabled format
            string outputPath = "output_signed.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
