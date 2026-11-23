// Title: Export a VBA project's digital certificate to a .cer file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export VBA project digital certificate to file | C# extract VBA macro certificate from Excel workbook with Aspose.Cells | How to save VBA digital signature as .cer using Aspose.Cells .NET | Retrieve VBA project digital signature when ExportDigitalSignature method is unavailable | Export VBA macro certificate with reflection in Aspose.Cells
// Tags: export VBA digital certificate Aspose.Cells | C# VBA macro certificate extraction | reflection call ExportDigitalSignature method | VbaProject certificate export to .cer | detect missing ExportDigitalSignature API | verify VBA project existence in workbook

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, checks for a VBA project, and uses reflection to invoke the ExportDigitalSignature method on the VbaProject to write the attached digital certificate to a .cer file, handling cases where the method is absent or no certificate is present.
    class ExportVbaCertificate
    {
        static void Main()
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "vba_certificate.cer";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that may contain a VBA project
                Workbook workbook = new Workbook(inputFile);

                // Check for the presence of a VBA project
                if (workbook.VbaProject != null)
                {
                    try
                    {
                        // Use reflection to call ExportDigitalSignature if it exists in the current Aspose.Cells version
                        MethodInfo exportMethod = workbook.VbaProject.GetType()
                            .GetMethod("ExportDigitalSignature", new[] { typeof(string) });

                        if (exportMethod != null)
                        {
                            exportMethod.Invoke(workbook.VbaProject, new object[] { outputFile });
                            Console.WriteLine($"Digital certificate exported to \"{outputFile}\".");
                        }
                        else
                        {
                            Console.WriteLine("ExportDigitalSignature method is not available in this Aspose.Cells version.");
                        }
                    }
                    catch (TargetInvocationException tie) when (tie.InnerException != null)
                    {
                        // ExportDigitalSignature throws if no certificate is attached
                        Console.WriteLine($"Unable to export digital certificate: {tie.InnerException.Message}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during certificate export: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("No VBA project found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
            }
        }
    }
}
