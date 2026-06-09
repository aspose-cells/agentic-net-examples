using System;
using Aspose.Cells;

namespace AsposeCellsLockOutsideRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the editable range (e.g., cells A1:C5)
            // Add it to the AllowEditRanges collection so it remains editable when the sheet is protected
            int rangeIndex = sheet.AllowEditRanges.Add("EditableRange", 0, 0, 4, 2);
            ProtectedRange editableRange = sheet.AllowEditRanges[rangeIndex];

            // (Optional) Set a password for the editable range
            // editableRange.Password = "rangePwd";

            // Protect the entire worksheet.
            // All cells are locked by default; the range added above is exempted from protection.
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("LockedOutsideRange.xlsx");
        }
    }
}