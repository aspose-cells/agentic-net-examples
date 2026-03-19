using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook (initially without a VBA project)
        Workbook workbook = new Workbook();

        // Save as a macro-enabled workbook to initialize the VBA project,
        // then reload it so the VbaProject object becomes available.
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);
        workbook = new Workbook("temp.xlsm");

        // Protect the VBA project, lock it for viewing, and set a password.
        // The first argument (true) locks the project for viewing.
        // The second argument is the password that will be required to view/unprotect.
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook as a macro-enabled template (XLTM) with the protected VBA project.
        workbook.Save("ProtectedTemplate.xltm", SaveFormat.Xltm);
    }
}