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

            // Add a registered reference to an Automation type library.
            // The libid is provided in the typical numeric format.
            string referenceName = "stdole";
            string libId = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation";

            vbaProject.References.AddRegisteredReference(referenceName, libId);

            // Save the workbook as a macro-enabled file to preserve the VBA project and its references
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}