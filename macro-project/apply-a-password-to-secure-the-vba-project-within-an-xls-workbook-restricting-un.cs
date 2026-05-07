using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook (XLSM format will be used when saving)
        Workbook workbook = new Workbook();

        // Protect the VBA project and lock it for viewing with a password
        // islockedForViewing = true ensures the project cannot be opened without the password
        workbook.VbaProject.Protect(true, "StrongPassword123");

        // Save the workbook as a macro‑enabled file
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}