using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectSharedWorkbookSample
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Enable sharing for the workbook
            wb.Settings.Shared = true;

            // Protect the shared workbook with a password
            wb.ProtectSharedWorkbook("myPassword");

            // Save the protected, shared workbook as an XLSX file
            wb.Save("ProtectedSharedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}