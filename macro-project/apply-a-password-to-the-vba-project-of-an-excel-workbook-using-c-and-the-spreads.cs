using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (initially empty)
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled file to ensure a VBA project container exists
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);

        // Reload the workbook so that the VBA project is initialized
        workbook = new Workbook("temp.xlsm");

        // Protect the VBA project with a password.
        // The first argument (islockedForViewing) is set to false because we only want to protect,
        // not lock the project for viewing.
        workbook.VbaProject.Protect(false, "MyVbaPassword");

        // Save the workbook in SpreadsheetML (XML) format.
        // The VBA project (now password‑protected) will be included in the saved file.
        workbook.Save("VbaProtected.xml", SaveFormat.SpreadsheetML);
    }
}