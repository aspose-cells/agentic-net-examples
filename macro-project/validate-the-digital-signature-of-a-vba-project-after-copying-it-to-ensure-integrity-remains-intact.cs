// Title: How to verify a VBA project's digital signature remains intact after copying an XLSM workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check if VBA project is still signed after saving as XLSM | verify VBA macro digital signature after copying workbook using Aspose.Cells .NET | how to ensure VBA project signature persists when duplicating an .xlsm file in C# | Aspose.Cells .NET validate signed VBA project in copied macro-enabled workbook | detect lost VBA signature after workbook copy with Aspose.Cells
// Tags: Aspose.Cells VBA signature integrity check | C# retain signed VBA macro in XLSM files | monitor VBA project signing status after file save | ensure macro digital signature persists with Aspose.Cells | use .NET Workbook API to validate VBA signing

using System;
using System.IO;
using Aspose.Cells;

namespace VbaSignatureValidatorApp
{
    // The example loads a source XLSM workbook, confirms it contains a signed VBA project, saves the workbook to a new file (copying the VBA project and its signature), reloads the copied file, and verifies that the VBA project remains signed, outputting a success or failure message.
    class VbaSignatureValidator
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsm";
                const string copiedPath = "copied.xlsm";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook containing the VBA project
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Ensure the workbook actually has a VBA project
                if (sourceWorkbook.VbaProject == null)
                {
                    Console.WriteLine("The source workbook does not contain a VBA project.");
                    return;
                }

                // Check if the source VBA project is signed
                bool sourceSigned = sourceWorkbook.VbaProject.IsSigned;

                // Save the workbook to a new file (this copies the VBA project and its signature)
                sourceWorkbook.Save(copiedPath, SaveFormat.Xlsm);

                // Verify the copied file was created
                if (!File.Exists(copiedPath))
                {
                    Console.WriteLine($"Failed to create copied workbook: {copiedPath}");
                    return;
                }

                // Reload the saved workbook to ensure the signature persisted
                Workbook reloadedWorkbook = new Workbook(copiedPath);

                // Validate that the VBA project is still signed after copying
                bool reloadedSigned = reloadedWorkbook.VbaProject != null && reloadedWorkbook.VbaProject.IsSigned;

                // Final integrity check
                if (sourceSigned && reloadedSigned)
                {
                    Console.WriteLine("VBA project digital signature is valid after copying.");
                }
                else
                {
                    Console.WriteLine("VBA project digital signature validation failed.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
