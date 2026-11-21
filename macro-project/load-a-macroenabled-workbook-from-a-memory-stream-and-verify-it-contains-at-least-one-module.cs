// Title: Load a macro‑enabled .xlsm workbook from a MemoryStream and check for VBA modules using Aspose.Cells for .NET
// Common Searches: how to read a .xlsm file from a byte array with Aspose.Cells in C# | verify that a macro-enabled workbook contains VBA modules using Aspose.Cells | load a macro-enabled Excel workbook from MemoryStream and inspect VbaProject in .NET | Aspose.Cells C# example to detect VBA modules in an .xlsm file | using LoadOptions to open .xlsm from a stream with Aspose.Cells
// Tags: load .xlsm workbook from memory stream Aspose.Cells | check VBA modules in Aspose.Cells workbook | VbaProject module count verification C# | LoadOptions Xlsx for macro-enabled files | read macro-enabled Excel bytes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads a macro‑enabled Excel file (.xlsm) into a byte array, creates a MemoryStream, and loads it with Aspose.Cells using LoadOptions. It then verifies that the workbook's VbaProject exists and contains at least one module, outputting the result while handling file‑not‑found and generic exceptions.
class MacroWorkbookVerifier
{
    static void Main()
    {
        try
        {
            // Obtain the macro‑enabled workbook bytes.
            byte[] workbookBytes = GetWorkbookBytes();

            // Load the workbook from a memory stream. LoadOptions without explicit format works for .xlsm.
            using (MemoryStream stream = new MemoryStream(workbookBytes))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(stream, loadOptions);

                // Verify that the workbook contains at least one VBA module.
                if (workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0)
                {
                    Console.WriteLine("The workbook contains VBA modules.");
                }
                else
                {
                    Console.WriteLine("The workbook does NOT contain any VBA modules.");
                }
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File not found: {ex.FileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retrieves the workbook bytes from a local file.
    static byte[] GetWorkbookBytes()
    {
        string path = "sample.xlsm";

        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The specified workbook file was not found.", path);
        }

        return File.ReadAllBytes(path);
    }
}
