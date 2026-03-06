using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Save as a temporary macro-enabled workbook to initialize the VBA project
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);

        // Load the temporary workbook to access the VBA project
        workbook = new Workbook(tempPath);

        // Delete the temporary file
        File.Delete(tempPath);

        // Protect the VBA project with a password (not locked for viewing)
        workbook.VbaProject.Protect(false, "vbaPassword123");

        // Save the workbook in ODS format
        workbook.Save("ProtectedVbaProject.ods", SaveFormat.Ods);
    }
}