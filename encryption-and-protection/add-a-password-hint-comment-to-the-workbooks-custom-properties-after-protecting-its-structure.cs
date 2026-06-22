using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's structure with a password
        workbook.Protect(ProtectionType.Structure, "SecretPassword123");

        // Add a custom document property that serves as a password hint
        // This property will be stored in the workbook and can be read later
        workbook.CustomDocumentProperties.Add("PasswordHint", "Your hint goes here");

        // Save the protected workbook
        workbook.Save("ProtectedWorkbookWithHint.xlsx");
    }
}