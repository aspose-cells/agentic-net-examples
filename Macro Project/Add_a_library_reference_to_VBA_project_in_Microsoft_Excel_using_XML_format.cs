using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReference
{
    static void Main()
    {
        // Create a new workbook (will become a macro‑enabled workbook when saved as Xlsm)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external type library (e.g., stdole)
        // name: logical name of the reference
        // libid: identifier of the Automation type library
        vbaProject.References.AddRegisteredReference(
            "stdole",
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

        // Save the workbook as a macro‑enabled file; the reference is stored in the VBA project XML
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}