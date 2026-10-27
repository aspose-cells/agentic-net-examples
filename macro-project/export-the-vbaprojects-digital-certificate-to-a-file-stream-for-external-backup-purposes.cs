// Title: Export a VBA project's digital certificate to a binary file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export VBA project binary from an XLSX workbook | How to backup a VBA digital certificate from Excel using Aspose.Cells .NET | Extract xl/vbaProject.bin from a temporary XLSM with Aspose.Cells in C# | Save VBA project to a file stream using Aspose.Cells for .NET | Retrieve VBA project certificate from workbook with Aspose.Cells API
// Tags: Aspose.Cells export VBA binary | C# read vbaProject.bin from XLSM zip | preserve VBA digital certificate using .NET | write VBA project to stream Aspose.Cells | load workbook access VbaProject API

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject and related classes

// The example checks for an input XLSX file, loads it with Aspose.Cells, verifies the presence of a VBA project, saves the workbook as a temporary XLSM, opens the XLSM as a zip archive, extracts the 'xl/vbaProject.bin' entry (which contains the VBA project and its digital certificate) to a specified output file, and cleans up the temporary file while handling errors.
class VbaCertificateExport
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "VbaProject.bin";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that may contain a VBA project.
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project of the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // Export the VBA project (including any digital certificate if present) to a binary file.
                // Since ExportVbaProject is not available in the referenced version, save as XLSM and extract the binary.
                string tempXlsm = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");

                try
                {
                    // Save workbook as XLSM (which contains the VBA project).
                    workbook.Save(tempXlsm, SaveFormat.Xlsm);

                    // Open the XLSM package and extract the vbaProject.bin entry.
                    using (ZipArchive archive = ZipFile.OpenRead(tempXlsm))
                    {
                        ZipArchiveEntry vbaEntry = archive.GetEntry("xl/vbaProject.bin");
                        if (vbaEntry != null)
                        {
                            vbaEntry.ExtractToFile(outputPath, true);
                            Console.WriteLine("VBA project exported successfully.");
                        }
                        else
                        {
                            Console.WriteLine("VBA project binary not found in the workbook.");
                        }
                    }
                }
                catch (Exception exportEx)
                {
                    Console.WriteLine($"Failed to export VBA project: {exportEx.Message}");
                }
                finally
                {
                    // Clean up temporary file.
                    if (File.Exists(tempXlsm))
                    {
                        try { File.Delete(tempXlsm); } catch { /* ignore cleanup errors */ }
                    }
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
