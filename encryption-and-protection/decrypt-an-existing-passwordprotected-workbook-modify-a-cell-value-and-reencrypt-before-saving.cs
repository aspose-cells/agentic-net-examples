using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing password‑protected workbook
            string inputFile = "protected.xlsx";

            // Path for the modified and re‑encrypted workbook
            string outputFile = "modified_protected.xlsx";

            // Password used to open and later protect the workbook
            string password = "mySecretPwd";

            // Load the workbook with the correct password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Modify a cell value (e.g., A1 on the first worksheet)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Updated value");

            // Re‑encrypt the workbook with the same password
            workbook.Settings.Password = password;

            // Save the workbook
            workbook.Save(outputFile);
        }
    }
}