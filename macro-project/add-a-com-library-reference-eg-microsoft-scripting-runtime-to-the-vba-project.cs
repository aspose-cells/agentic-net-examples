using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (will be saved as macro-enabled .xlsm)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to the Microsoft Scripting Runtime COM library
            // Name: "Scripting"
            // LibID format: "*\\G{<GUID>}#<Version>#0#<Path>#<Description>"
            vbaProject.References.AddRegisteredReference(
                "Scripting",
                "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\system32\\scrrun.dll#Microsoft Scripting Runtime"
            );

            // Save the workbook as a macro-enabled file
            workbook.Save("WorkbookWithScriptingReference.xlsm", SaveFormat.Xlsm);
        }
    }
}