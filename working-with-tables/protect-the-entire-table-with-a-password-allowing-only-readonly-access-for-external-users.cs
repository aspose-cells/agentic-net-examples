using System;
using Aspose.Cells;

namespace AsposeCellsReadOnlyProtection
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Add some data to demonstrate the protected area
            sheet.Cells["A1"].PutValue("Read‑only data");
            sheet.Cells["B2"].PutValue(12345);

            // Protect the entire worksheet with a password.
            // ProtectionType.All protects all aspects (structure, objects, etc.).
            // The third parameter (oldPassword) is null because the sheet is not previously protected.
            sheet.Protect(ProtectionType.All, "SecureReadOnlyPassword", null);

            // Save the workbook; external users will be able to open it read‑only only.
            workbook.Save("ReadOnlyProtectedWorkbook.xlsx");
        }
    }
}