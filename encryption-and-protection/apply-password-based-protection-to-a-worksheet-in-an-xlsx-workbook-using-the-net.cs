using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password and all protection types
            // The third parameter (oldPassword) is null because the sheet is not previously protected
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Save the protected workbook to an XLSX file
            workbook.Save("ProtectedWorksheet.xlsx", SaveFormat.Xlsx);
        }
    }
}