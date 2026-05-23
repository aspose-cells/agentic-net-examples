using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the editable area (example: B2:D10) and give it a name
        // Parameters: name, startRow, startColumn, endRow, endColumn (zero‑based indices)
        int editRangeIndex = worksheet.AllowEditRanges.Add("EditableSection", 1, 1, 9, 3);
        // Optional: set a password for the editable range
        worksheet.AllowEditRanges[editRangeIndex].Password = "editpwd";

        // Protect the entire worksheet (all protection types) with a sheet password
        worksheet.Protect(ProtectionType.All, "sheetpwd", null);

        // Save the workbook
        workbook.Save("ProtectedWithEditableSection.xlsx");
    }
}