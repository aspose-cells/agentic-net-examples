using System;
using Aspose.Cells;
using System.IO;

class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Ensure a VBA project exists by saving as a macro-enabled workbook and reloading it
        string tempPath = "temp.xlsm";
        wb.Save(tempPath);
        wb = new Workbook(tempPath);
        File.Delete(tempPath); // clean up temporary file

        // Protect the VBA project with a password (not locked for viewing)
        wb.VbaProject.Protect(false, "VbaPassword123");

        // Save the workbook in XLS format
        wb.Save("ProtectedVbaProject.xls");
    }
}