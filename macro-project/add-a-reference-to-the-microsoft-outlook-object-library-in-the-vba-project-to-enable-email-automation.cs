// Title: Aspose.Cells for .NET cannot add Microsoft Outlook Object Library reference to a VBA project in an Xlsm workbook
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel file, detects whether a VBA project is present, and logs that Outlook Object Library references cannot be added via the API. | Describe the manual steps to insert the Microsoft Outlook Object Library reference into a macro‑enabled workbook after it has been saved with Aspose.Cells.
// Common Searches: aspnet add outlook object library to vba project using aspose.cells | why Aspose.Cells cannot programmatically insert VBA references | how to include Outlook reference in xlsm file created by Aspose.Cells | saving macro-enabled workbook with existing VBA project aspose.cells limitation
// Tags: aspose.cells vba reference limitation | outlook object library reference in xlsm | c# load workbook check vba project | save macro-enabled workbook aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads or creates an Excel workbook, checks for an existing VBA project, reports that adding VBA references such as the Microsoft Outlook Object Library cannot be performed through the Aspose.Cells API, and then saves the file as a macro‑enabled Xlsm workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsm";

            // Load existing workbook if it exists; otherwise create a new empty workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new empty workbook
            }

            // Check whether the workbook already contains a VBA project.
            if (workbook.VbaProject != null)
            {
                // Note: Adding VBA references programmatically is not supported in the current API.
                Console.WriteLine("VBA project detected (references cannot be added via API).");
            }
            else
            {
                Console.WriteLine("No VBA project present in the workbook.");
            }

            // Save as a macro‑enabled workbook to preserve any VBA project.
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
