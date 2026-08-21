// Title: C# – Add a Registered Automation Library Reference to a VBA Project with Aspose.Cells
// Description: Creates a macro‑enabled workbook, accesses its VbaProject, adds a registered reference to the stdole type library via VbaProject.References.AddRegisteredReference, and saves the file as an XLSM workbook.
// Keywords: Aspose.Cells VBA reference C# | VbaProject.References.AddRegisteredReference | add stdole library Aspose.Cells | macro‑enabled workbook automation | C# Aspose.Cells VBA automation reference
// Common Searches: how to add a registered type library to a VBA project using Aspose.Cells .NET | Aspose.Cells C# add stdole reference to XLSM workbook | VbaProject References AddRegisteredReference example | save workbook with VBA references Aspose.Cells
// Developer Intent: Programmatically attach a registered automation type‑library (e.g., stdole) to a workbook’s VBA project and save it as a macro‑enabled file.
// Use Cases: Enable generated macros to use OLE Automation objects by adding the stdole reference. | Prepare a batch of workbooks with required COM library references before distribution. | Standardize VBA environments across multiple projects by automating reference insertion.
// AI Prompts: Generate C# code that adds a registered reference to a VBA project using Aspose.Cells, including robust error handling. | Show how to add multiple registered references (such as stdole and Microsoft Office) to a workbook’s VBA project with Aspose.Cells. | Explain how to confirm that a registered reference was successfully added after saving the workbook as an XLSM file.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    // Creates a macro‑enabled workbook, accesses its VbaProject, adds a registered reference to the stdole type library via VbaProject.References.AddRegisteredReference, and saves the file as an XLSM workbook.
    public class AddRegisteredReferenceExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (macro-enabled format will be used when saving)
                Workbook workbook = new Workbook();

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Add a registered reference to an Automation type library (e.g., stdole)
                vbaProject.References.AddRegisteredReference(
                    "stdole",
                    "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

                // Save the workbook as a macro-enabled file
                workbook.Save("AddedRegisteredReference.xlsm", SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddRegisteredReferenceExample.Run();
        }
    }
}
