using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddComReference
{
    static void Main()
    {
        // Create a new workbook (macro-enabled)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Define the COM library reference (Automation type library)
        // Example: Microsoft Scripting Runtime (scrrun.dll)
        string referenceName = "Scripting";
        string libid = "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Scripting Runtime";

        // Add the reference to the VBA project
        vbaProject.References.AddRegisteredReference(referenceName, libid);

        // Save the workbook as a macro-enabled file
        workbook.Save("WorkbookWithComReference.xlsm", SaveFormat.Xlsm);
    }
}