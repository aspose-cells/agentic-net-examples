using System;
using Aspose.Cells;

namespace AsposeCellsPasswordSaveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set the password that will protect the workbook when saved
            // This utilizes the WorkbookSettings.Password property (save‑with‑password rule)
            workbook.Settings.Password = "MySecretPassword";

            // Create OOXML save options (creation rule for OoxmlSaveOptions)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            // Example of setting an additional option (optional)
            saveOptions.CompressionType = OoxmlCompressionType.Level6;

            // Save the workbook as XLSX with the password protection enabled
            // This follows the save rule: workbook.Save(filePath, saveOptions)
            workbook.Save("ProtectedWorkbook.xlsx", saveOptions);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook saved as 'ProtectedWorkbook.xlsx' with password protection.");
        }
    }
}