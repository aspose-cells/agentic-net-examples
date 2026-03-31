using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.IO;

class ApplyVbaPassword
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled file to create a VBA project, then reload it
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);
        File.Delete(tempFile); // clean up temporary file

        // Protect the VBA project, lock it for viewing, and set a password
        workbook.VbaProject.Protect(true, "MyVbaPassword");

        // Save the workbook with the protected VBA project
        workbook.Save("VbaProtectedWorkbook.xlsm", SaveFormat.Xlsm);
    }
}