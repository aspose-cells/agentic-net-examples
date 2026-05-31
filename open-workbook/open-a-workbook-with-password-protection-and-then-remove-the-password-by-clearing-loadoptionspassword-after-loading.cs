using Aspose.Cells;

class RemoveWorkbookPassword
{
    static void Main()
    {
        // Path to the password‑protected workbook
        string inputPath = "protected.xlsx";
        string outputPath = "unprotected.xlsx";

        // Create LoadOptions and set the password needed to open the workbook
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "test";

        // Load the workbook using the provided password
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Clear the password from LoadOptions after loading (demonstrates removal)
        loadOptions.Password = null;

        // Remove any workbook encryption password
        workbook.Settings.Password = null;

        // Save the workbook without password protection
        workbook.Save(outputPath);
    }
}