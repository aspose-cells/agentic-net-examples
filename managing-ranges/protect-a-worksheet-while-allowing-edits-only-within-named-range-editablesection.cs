using Aspose.Cells;

class ProtectWorksheetExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the editable area (example: B2:D10)
        int startRow = 1;      // Row 2 (zero‑based)
        int startColumn = 1;   // Column B
        int endRow = 9;        // Row 10
        int endColumn = 3;     // Column D

        // Add an allow‑edit range named "EditableSection"
        int rangeIndex = sheet.AllowEditRanges.Add("EditableSection", startRow, startColumn, endRow, endColumn);
        // Optional: set a password for the range
        // sheet.AllowEditRanges[rangeIndex].Password = "rangePassword";

        // Protect the worksheet (all protection types) with a password
        sheet.Protect(ProtectionType.All, "sheetPassword", null);

        // Save the workbook
        workbook.Save("ProtectedWithEditableSection.xlsx");
    }
}