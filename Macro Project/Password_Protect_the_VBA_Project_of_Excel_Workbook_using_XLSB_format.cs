using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (XLSB supports macros)
        Workbook workbook = new Workbook();

        // Protect the VBA project with a password (do not lock for viewing)
        workbook.VbaProject.Protect(false, "MySecretPassword");

        // Save the workbook in XLSB format
        workbook.Save("ProtectedVbaProject.xlsb", SaveFormat.Xlsb);
    }
}