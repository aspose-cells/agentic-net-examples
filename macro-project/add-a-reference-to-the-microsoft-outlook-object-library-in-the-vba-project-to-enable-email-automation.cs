// Title: Add Outlook Object Library Reference to an Excel VBA Project with Aspose.Cells (C#)
// Description: Demonstrates how to create a new workbook with Aspose.Cells, access its VBA project, set UTF‑8 encoding, and programmatically add a registered reference to the Microsoft Outlook 16.0 Object Library. The workbook is saved as a macro‑enabled XLSM file, ready for Outlook automation from VBA macros.
// Keywords: Aspose.Cells VBA reference | C# add Outlook library | Excel macro‑enabled workbook | Add Outlook reference programmatically | VbaProject References AddRegisteredReference | Outlook Object Library GUID | Excel automation with Outlook | Save XLSM with Aspose.Cells
// Common Searches: how to add outlook reference to excel vba using aspocells | c# add microsoft outlook object library to vba project | programmatically insert outlook reference in xlsm file | aspocells add registered reference outlook | create macro enabled workbook with outlook automation
// Developer Intent: Insert the Microsoft Outlook Object Library into a workbook’s VBA project so VBA macros can control Outlook.
// Use Cases: Generate a template workbook that can send emails via Outlook without manual reference setup. | Automate report distribution where VBA code drafts and dispatches Outlook messages. | Prepare a pre‑configured XLSM file for end‑users who need Outlook integration in Excel macros.
// AI Prompts: Write C# code using Aspose.Cells to add the Outlook 16.0 Object Library as a registered reference in a VBA project and save the file as .xlsm. | Explain how to locate the Outlook library GUID and use VbaProject.References.AddRegisteredReference to embed it in an Excel workbook. | Show how to verify that the Outlook reference was added correctly to the generated VBA project.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaOutlookReference
{
    // Demonstrates how to create a new workbook with Aspose.Cells, access its VBA project, set UTF‑8 encoding, and programmatically add a registered reference to the Microsoft Outlook 16.0 Object Library. The workbook is saved as a macro‑enabled XLSM file, ready for Outlook automation from VBA macros.
    public class AddOutlookReference
    {
        public static void Run()
        {
            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Get (or lazily create) the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Set the encoding for the VBA project
                vbaProject.Encoding = Encoding.UTF8;

                // Add a reference to the Microsoft Outlook Object Library
                string outlookLibId = "*\\G{00062FFF-0000-0000-C000-000000000046}#9.0#0#C:\\Program Files\\Microsoft Office\\Office16\\MSOUTL.OLB#Microsoft Outlook 16.0 Object Library";
                vbaProject.References.AddRegisteredReference("Outlook", outlookLibId);

                // Define output path
                string outputPath = "WorkbookWithOutlookReference.xlsm";

                // Ensure the output directory exists (if any)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a macro‑enabled file
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddOutlookReference.Run();
        }
    }
}
