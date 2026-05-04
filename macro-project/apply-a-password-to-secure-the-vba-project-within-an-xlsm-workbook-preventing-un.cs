using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtection
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (initially without VBA project)
            Workbook workbook = new Workbook();

            // Save as macro-enabled workbook to initialize the VBA project
            workbook.Save("temp.xlsm", SaveFormat.Xlsm);

            // Reload the workbook so that the VBA project is available
            workbook = new Workbook("temp.xlsm");

            // Apply protection to the VBA project and lock it for viewing
            // islockedForViewing = true, password = "MySecretPassword"
            workbook.VbaProject.Protect(true, "MySecretPassword");

            // Save the protected workbook
            workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);

            // Optional: clean up temporary file
            System.IO.File.Delete("temp.xlsm");
        }
    }
}