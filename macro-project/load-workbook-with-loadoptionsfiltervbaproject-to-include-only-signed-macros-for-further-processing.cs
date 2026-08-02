// Title: Load Only Signed VBA Macros with Aspose.Cells for .NET using LoadOptions.FilterVbaProject
// Description: Demonstrates how to open an XLSM workbook with Aspose.Cells, filter the VBA project to include only digitally signed macros via LoadOptions.FilterVbaProject, verify the signature with VbaProject.IsSigned and VbaProject.IsValidSigned, and optionally save the workbook.
// Keywords: Aspose.Cells LoadOptions.FilterVbaProject | load signed VBA macros .NET | check VBA project signature Aspose | filter unsigned macros Excel | Aspose.Cells VBA validation | C# Aspose.Cells signed macro example | GitHub Aspose.Cells VBA sample
// Common Searches: Aspose.Cells load only signed VBA macros | filter VBA project by digital signature .NET | how to verify signed VBA project with Aspose.Cells | C# example for loading signed macros in Excel | Aspose.Cells LoadOptions.FilterVbaProject usage
// Developer Intent: Open an Excel file and process it only when the embedded VBA project is digitally signed.
// Use Cases: Ensure macro security by processing workbooks that contain a trusted signed VBA project. | Skip or reject files with unsigned or missing VBA projects before automation. | Log signature status (IsSigned, IsValidSigned) and preserve the signed macro when saving the workbook.
// AI Prompts: Write C# code that uses LoadOptions.FilterVbaProject to load only signed VBA macros from an XLSM file with Aspose.Cells. | Show how to raise a custom exception if workbook.VbaProject.IsSigned returns false. | Create a logging snippet that records VbaProject.IsSigned and VbaProject.IsValidSigned, then saves the workbook conditionally.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to open an XLSM workbook with Aspose.Cells, filter the VBA project to include only digitally signed macros via LoadOptions.FilterVbaProject, verify the signature with VbaProject.IsSigned and VbaProject.IsValidSigned, and optionally save the workbook.
class LoadSignedVbaWorkbook
{
    static void Main()
    {
        string inputPath = "SignedWorkbook.xlsm";
        string outputPath = "ProcessedWorkbook.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook; VBA project is loaded automatically if present
            Workbook workbook = new Workbook(inputPath);

            // Check if a VBA project is present and whether it is signed
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine($"VBA project is signed. Signature valid: {workbook.VbaProject.IsValidSigned}");
            }
            else
            {
                Console.WriteLine("No signed VBA project loaded.");
            }

            // Save the workbook after processing (optional)
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
