// Title: Validate VBA Project Digital Signature in an .xlsm Workbook with Aspose.Cells for .NET (C#)
// Description: C# example that checks if a macro‑enabled workbook exists, loads it with Aspose.Cells, accesses its VbaProject, determines whether the VBA project is signed, validates the signature using IsValidSigned, and prints the result while handling errors.
// Keywords: Aspose.Cells VBA signature validation | C# IsValidSigned example | Workbook.VbaProject signed macro check | digital signature verification .xlsm | macro security Aspose.Cells | validate signed VBA project .NET
// Common Searches: how to verify VBA project signature with Aspose.Cells C# | C# code to check if Excel macro is signed | Workbook.VbaProject.IsValidSigned usage | validate signed .xlsm file using Aspose.Cells | Aspose.Cells verify macro digital signature
// Developer Intent: Determine whether a workbook's VBA project is signed and capture the validity of its digital signature using Aspose.Cells.
// Use Cases: Enforce macro security policies by allowing only workbooks with a valid VBA signature to run. | Automate trust assessment of incoming .xlsm files in a document‑processing pipeline and log signature status. | Generate compliance reports that include signature validation outcomes for all processed macro‑enabled workbooks.
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsm file, checks if its VBA project is signed, validates the signature, and returns a boolean indicating validity. | Create a reusable method that accepts a file path, verifies the presence of a signed VBA project, validates the signature, and returns a detailed result object with status messages and error handling. | Provide best practices for handling exceptions and unsigned projects when validating VBA signatures with Aspose.Cells in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Namespace containing VbaProject

// C# example that checks if a macro‑enabled workbook exists, loads it with Aspose.Cells, accesses its VbaProject, determines whether the VBA project is signed, validates the signature using IsValidSigned, and prints the result while handling errors.
class ValidateVbaSignature
{
    static void Main()
    {
        const string filePath = "sample.xlsm";

        // Ensure the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load a workbook that contains a VBA project
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Capture the result of the signature validation
                bool isSignatureValid = vbaProject.IsValidSigned;

                // Output the validation result
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature valid: " + isSignatureValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or not present.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
