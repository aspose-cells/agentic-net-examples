using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtection
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // (Optional) Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Protect the VBA project with a password.
            // The first argument (false) means the project is not locked for viewing.
            // The second argument is the password that will be required to unprotect the VBA project.
            workbook.VbaProject.Protect(false, "VbaProjectPassword123");

            // Save the workbook as an ODS file.
            // The VBA project password is retained in the saved file.
            workbook.Save("VbaProtectedWorkbook.ods", SaveFormat.Ods);
        }
    }
}