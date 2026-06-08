using Aspose.Cells;

class ProtectWorkbookProperties
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set default document properties (author and title)
        workbook.BuiltInDocumentProperties.Author = "Original Author";
        workbook.BuiltInDocumentProperties.Title = "Original Title";

        // Apply write protection to prevent changes to these properties
        workbook.Settings.WriteProtection.Password = "protect123";   // password required to modify
        workbook.Settings.WriteProtection.Author = "Original Author"; // author stored with protection
        workbook.Settings.WriteProtection.RecommendReadOnly = true;   // optional: recommend read‑only

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}