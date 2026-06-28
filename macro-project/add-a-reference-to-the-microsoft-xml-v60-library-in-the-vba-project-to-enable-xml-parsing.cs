using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default is .xlsx)
            Workbook workbook = new Workbook();

            // Access the VBA project (automatically created for macro-enabled workbooks)
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to Microsoft XML, v6.0 (MSXML6) as a registered type library
            // Libid format: *\G{<GUID>}#<Version>#0#<Path>#<Description>
            // GUID for MSXML6: 88D969C0-F192-11D4-A65F-0040963251E5 (example)
            // Adjust the path if MSXML6 is located elsewhere.
            string name = "MSXML2";
            string libid = "*\\G{88D969C0-F192-11D4-A65F-0040963251E5}#6.0#0#C:\\Windows\\System32\\msxml6.dll#Microsoft XML, v6.0";

            vbaProject.References.AddRegisteredReference(name, libid);

            // Save the workbook as a macro-enabled file to retain the VBA project and its references
            workbook.Save("WorkbookWithMSXMLReference.xlsm", SaveFormat.Xlsm);
        }
    }
}