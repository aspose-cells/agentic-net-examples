using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook. VBA project is automatically created for macro-enabled formats.
        Workbook workbook = new Workbook();

        // Get the VBA project associated with the workbook.
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an Automation type library (e.g., stdole).
        // Parameters: reference name and its LIBID.
        vbaProject.References.AddRegisteredReference(
            "stdole",
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

        // Save the workbook as a macro‑enabled template (XLTM format).
        workbook.Save("WorkbookWithReference.xltm", SaveFormat.Xltm);
    }
}