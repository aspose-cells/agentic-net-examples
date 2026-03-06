using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply password protection (encryption)
        workbook.Settings.Password = "mySecurePassword";

        // Optionally specify encryption algorithm and key length
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Create ODS save options and set desired properties
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice; // optional

        // Save the encrypted workbook as an ODS file
        workbook.Save("encrypted_output.ods", saveOptions);
    }
}