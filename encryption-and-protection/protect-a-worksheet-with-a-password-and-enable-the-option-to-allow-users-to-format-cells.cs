using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectWorksheetAllowFormattingCell
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the protection object for the worksheet
            Protection protection = sheet.Protection;

            // Allow users to format cells even when the sheet is protected
            protection.AllowFormattingCell = true;

            // Set the password that will protect the worksheet
            string password = "MySecretPassword";
            protection.Password = password;

            // Apply protection to the worksheet with the specified password
            // Using ProtectionType.All to protect all aspects except those explicitly allowed
            sheet.Protect(ProtectionType.All, password, null);

            // Define output file path
            string outputPath = "ProtectedWorksheet.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}