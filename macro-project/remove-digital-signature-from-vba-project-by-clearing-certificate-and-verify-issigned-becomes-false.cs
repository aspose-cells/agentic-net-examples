// Title: C# – Remove Digital Signature from VBA Project and Verify IsSigned with Aspose.Cells
// Description: Loads a signed .xlsm workbook, checks the VBA project's IsSigned flag, removes the macro (which clears the digital signature), saves the file as .xlsx, reloads it, and confirms that IsSigned is false using Aspose.Cells for .NET.
// Keywords: Aspose.Cells remove VBA signature | clear VBA digital certificate .NET | IsSigned property Aspose.Cells | remove macro workbook C# | verify unsigned VBA project | Aspose.Cells VBA project handling | C# workbook digital signature removal
// Common Searches: how to delete digital signature from a signed VBA project using Aspose.Cells | C# remove VBA macro and clear certificate Aspose.Cells | check IsSigned flag after removing VBA project | Aspose.Cells remove macro unsigned workbook | verify VBA project is unsigned after RemoveMacro
// Developer Intent: Strip the digital signature from a signed VBA project and ensure the workbook reports the VBA project as unsigned.
// Use Cases: Automate compliance by batch‑processing .xlsm files to remove macros and their signatures before distribution. | Validate that a workbook no longer contains a signed VBA project after macro removal. | Convert signed macro‑enabled workbooks to unsigned .xlsx files for environments that prohibit macros.
// AI Prompts: Generate C# code with Aspose.Cells that removes a VBA project's digital signature and verifies IsSigned becomes false. | Explain why the RemoveMacro method also clears the certificate of a signed VBA project in Aspose.Cells. | Provide robust error‑handling patterns for removing VBA signatures when the workbook may lack a VBA project.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm workbook, checks the VBA project's IsSigned flag, removes the macro (which clears the digital signature), saves the file as .xlsx, reloads it, and confirms that IsSigned is false using Aspose.Cells for .NET.
    public class RemoveVbaSignatureDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Path to a macro‑enabled workbook that already contains a signed VBA project
            string signedPath = "SignedVbaWorkbook.xlsm";

            // Verify the source file exists
            if (!File.Exists(signedPath))
            {
                Console.WriteLine($"Source file not found: {signedPath}");
                return;
            }

            try
            {
                // Load the signed workbook
                Workbook workbook = new Workbook(signedPath);

                // Check if a VBA project exists before accessing its properties
                bool isSigned = workbook.VbaProject != null && workbook.VbaProject.IsSigned;
                Console.WriteLine("Before removal - VBA project signed: " + isSigned);

                // Remove the VBA project (macro) from the workbook.
                // This also clears any digital signature associated with the VBA project.
                workbook.RemoveMacro();

                // Save the workbook without the VBA project (and without its signature)
                string unsignedPath = "UnsignedVbaWorkbook.xlsx";
                workbook.Save(unsignedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved without VBA project: {unsignedPath}");

                // Reload the saved file to confirm the signature state
                Workbook reloaded = new Workbook(unsignedPath);
                bool isSignedAfter = reloaded.VbaProject != null && reloaded.VbaProject.IsSigned;
                Console.WriteLine("After removal - VBA project signed: " + isSignedAfter);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine("File not found: " + fnfEx.FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error processing workbook: " + ex.Message);
            }
        }
    }
}
