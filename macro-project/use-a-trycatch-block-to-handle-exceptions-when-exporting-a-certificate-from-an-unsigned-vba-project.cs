// Title: Export VBA Project Certificate with Try‑Catch Error Handling Using Aspose.Cells for .NET
// Description: Loads an .xlsm workbook with Aspose.Cells, checks file existence, uses nested try‑catch blocks to manage loading errors, unsigned VBA projects, and unexpected exceptions, and writes the raw certificate bytes to a .cer file when the project is signed.
// Keywords: Aspose.Cells export VBA certificate | C# VBA project certificate extraction | try catch Aspose.Cells loading error | unsigned VBA project handling | export .cer from Excel macro | VbaProject.IsSigned | CellsException handling
// Common Searches: how to export a VBA project certificate using Aspose.Cells C# | C# try‑catch example for extracting VBA certificate from .xlsm | what exception is thrown for an unsigned VBA project in Aspose.Cells | save VBA certificate bytes to a .cer file in .NET | Aspose.Cells error codes when loading a workbook
// Developer Intent: Retrieve and save the signing certificate of a VBA project while gracefully handling missing files, unsigned projects, and Aspose.Cells errors.
// Use Cases: Audit security by exporting the certificate of a signed VBA macro to a .cer file. | Prevent application crashes by informing users when a VBA project is unsigned. | Log detailed Aspose.Cells error codes for troubleshooting workbook load failures.
// AI Prompts: Write a C# method that returns the VBA project's certificate as a byte array using Aspose.Cells, with proper exception handling for unsigned projects and load errors. | Generate unit tests for ExportVbaCertificate covering signed, unsigned, missing file, and CellsException scenarios. | Refactor the sample to use async file I/O while preserving the same try‑catch structure and error messages.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads an .xlsm workbook with Aspose.Cells, checks file existence, uses nested try‑catch blocks to manage loading errors, unsigned VBA projects, and unexpected exceptions, and writes the raw certificate bytes to a .cer file when the project is signed.
class ExportVbaCertificate
{
    public static void Run()
    {
        // Path to the workbook that may contain a VBA project
        string workbookPath = "UnsignedWorkbook.xlsm";

        // Verify that the input file exists to prevent FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(workbookPath);
        }
        catch (CellsException cex)
        {
            Console.WriteLine("Aspose.Cells error while loading workbook: " + cex.Message);
            Console.WriteLine("Error code: " + cex.Code);
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error while loading workbook: " + ex.Message);
            return;
        }

        VbaProject vbaProject = workbook.VbaProject;

        try
        {
            // Verify that the VBA project is signed before accessing the certificate
            if (!vbaProject.IsSigned)
            {
                throw new InvalidOperationException("The VBA project is not signed; certificate data is unavailable.");
            }

            // Retrieve the raw certificate data
            byte[] certData = vbaProject.CertRawData;

            // Ensure data exists before writing to a file
            if (certData != null && certData.Length > 0)
            {
                File.WriteAllBytes("ExportedVbaCertificate.cer", certData);
                Console.WriteLine("Certificate exported successfully to ExportedVbaCertificate.cer");
            }
            else
            {
                Console.WriteLine("Certificate data is empty.");
            }
        }
        catch (InvalidOperationException ex)
        {
            // Handle the case where the VBA project is unsigned
            Console.WriteLine("Operation error: " + ex.Message);
        }
        catch (CellsException cex)
        {
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine("Aspose.Cells error: " + cex.Message);
            Console.WriteLine("Error code: " + cex.Code);
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }

    // Entry point required for compilation
    static void Main(string[] args)
    {
        Run();
    }
}
