using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro-enabled format will be used when saving)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Define the COM library reference details
        // Example: reference to the OLE Automation type library (stdole)
        // The libid string follows the pattern:
        // "*\\G{<GUID>}#<Version>#0#<FullPath>#<Description>"
        string referenceName = "stdole";
        string libid = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation";

        // Add the COM library reference to the VBA project
        vbaProject.References.AddRegisteredReference(referenceName, libid);

        // Save the workbook as a macro-enabled file to retain the VBA project and its references
        workbook.Save("WorkbookWithComReference.xlsm", SaveFormat.Xlsm);
    }
}