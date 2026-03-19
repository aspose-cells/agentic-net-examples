using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceToOds
{
    static void Main()
    {
        // Create a new workbook (ODS format)
        Workbook workbook = new Workbook();

        // Access the VBA project (created automatically for macro-enabled workbooks)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a registered reference to an external type library (example: stdole)
        vbaProject.References.AddRegisteredReference(
            "stdole",
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

        // Save the workbook as an ODS file
        workbook.Save("WorkbookWithVbaReference.ods", SaveFormat.Ods);
    }
}