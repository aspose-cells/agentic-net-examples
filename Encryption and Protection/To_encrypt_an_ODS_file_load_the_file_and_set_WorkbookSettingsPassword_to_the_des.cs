using Aspose.Cells;
using Aspose.Cells.Ods;

class EncryptOds
{
    static void Main()
    {
        // Load the existing ODS file (no password required for an unprotected file)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook("input.ods", loadOptions);

        // Set the password that will encrypt the workbook when saved
        workbook.Settings.Password = "Secret123";

        // Create ODS save options (optional, can be omitted)
        OdsSaveOptions saveOptions = new OdsSaveOptions();

        // Save the workbook as an encrypted ODS file
        workbook.Save("encrypted_output.ods", saveOptions);
    }
}