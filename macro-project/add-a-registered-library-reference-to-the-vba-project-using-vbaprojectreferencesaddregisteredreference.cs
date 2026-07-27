using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a registered reference to the VBA project (example: stdole library)
            vbaProject.References.AddRegisteredReference(
                "stdole",
                "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

            // Save the workbook as a macro-enabled file
            workbook.Save("VbaProjectWithRegisteredReference.xlsm", SaveFormat.Xlsm);
        }
    }
}