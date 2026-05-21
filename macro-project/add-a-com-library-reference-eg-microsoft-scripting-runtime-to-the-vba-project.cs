using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled format will be used on save)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Define the LibID for the Microsoft Scripting Runtime (scrrun.dll)
            // The format is: *\G{<GUID>}#<Version>#0#<FullPath>#<Description>
            string scriptingLibId = "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime";

            // Add the COM library reference to the VBA project
            vbaProject.References.AddRegisteredReference("Scripting", scriptingLibId);

            // Save the workbook as a macro-enabled file (XLSM)
            workbook.Save("WorkbookWithScriptingReference.xlsm", SaveFormat.Xlsm);
        }
    }
}