using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook (initially without a VBA project)
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled file and reload to create an empty VBA project
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);
        System.IO.File.Delete(tempFile);

        // Protect the VBA project and lock it for viewing with a password
        // islockedForViewing = true ensures the code cannot be viewed without the password
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook containing the protected VBA project
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}