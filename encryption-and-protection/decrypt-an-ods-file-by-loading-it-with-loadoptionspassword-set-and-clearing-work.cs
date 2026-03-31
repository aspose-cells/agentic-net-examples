using System;
using Aspose.Cells;

namespace DecryptOdsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the password‑protected ODS file
            string odsPath = "protected.ods";

            // Path for the decrypted XLSX output
            string xlsxPath = "decrypted.xlsx";

            // Create OdsLoadOptions and set the password required to open the ODS file
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            loadOptions.Password = "your_password";

            // Load the encrypted ODS workbook using the load options
            Workbook workbook = new Workbook(odsPath, loadOptions);

            // Clear the workbook password to remove encryption
            workbook.Settings.Password = null;

            // Save the workbook as an unprotected XLSX file
            workbook.Save(xlsxPath, SaveFormat.Xlsx);
        }
    }
}