// Title: Export VBA Project Certificate to Base64 with Aspose.Cells for .NET (C#)
// Description: Loads a workbook that contains a signed VBA project, extracts the certificate's raw bytes via Aspose.Cells VbaProject.CertRawData, converts them to a Base64 string, outputs the result, and saves the workbook unchanged.
// Keywords: Aspose.Cells VBA certificate export | C# convert VBA cert to Base64 | retrieve CertRawData Aspose.Cells | signed VBA project certificate .NET | export VBA project certificate
// Common Searches: how to get VBA project certificate with Aspose.Cells | convert VBA certificate to Base64 in C# | Aspose.Cells read signed VBA project certificate | export VBA cert raw data .NET | Base64 string from VBA project certificate
// Developer Intent: Extract the signed VBA project's certificate from a workbook and represent it as a Base64 string.
// Use Cases: Store the Base64 certificate in configuration files for macro integrity checks. | Log the certificate value for compliance auditing of signed workbooks. | Pass the certificate to deployment scripts that need to re‑sign macros on other machines.
// AI Prompts: Generate C# code using Aspose.Cells that reads a signed VBA project's CertRawData, converts it to Base64, and handles missing or unsigned projects gracefully. | Create a reusable method that returns the Base64 representation of a workbook's VBA certificate and logs appropriate messages for error conditions. | Write error‑handling logic for exporting a VBA certificate to Base64 while ensuring the workbook is saved without modifications.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a workbook that contains a signed VBA project, extracts the certificate's raw bytes via Aspose.Cells VbaProject.CertRawData, converts them to a Base64 string, outputs the result, and saves the workbook unchanged.
public class ExportVbaCertificate
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Path to the workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";

        // Verify the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed and certificate data exists
            if (vbaProject != null && vbaProject.IsSigned)
            {
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Convert the raw certificate bytes to a Base64 string
                    string base64Cert = Convert.ToBase64String(certData);

                    // Output the Base64 string (can be stored in configuration)
                    Console.WriteLine("VBA Project Certificate (Base64):");
                    Console.WriteLine(base64Cert);
                }
                else
                {
                    Console.WriteLine("Certificate raw data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed or not present.");
            }

            // Save the workbook (no modifications made, just demonstrating save lifecycle)
            string outputPath = "ExportVbaCertificate_Output.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
