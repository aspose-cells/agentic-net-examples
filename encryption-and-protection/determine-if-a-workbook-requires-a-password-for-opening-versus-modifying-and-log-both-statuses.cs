using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the demo workbook
            string filePath = "PasswordDemo.xlsx";

            // -------------------------------------------------
            // Create a new workbook and apply passwords
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Password required to open the workbook
            workbook.Settings.Password = "open123";

            // Password required to modify (write-protect) the workbook
            workbook.Settings.WriteProtection.Password = "modify123";

            // Save the workbook with the applied protections
            workbook.Save(filePath);
            workbook.Dispose();

            // -------------------------------------------------
            // Load the workbook using the opening password
            // -------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "open123"
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // -------------------------------------------------
            // Determine password requirements
            // -------------------------------------------------
            // Indicates if a password is needed to open the file
            bool requiresOpenPassword = loadedWorkbook.Settings.IsEncrypted;

            // Indicates if the workbook is write-protected (requires a password to modify)
            bool requiresModifyPassword = loadedWorkbook.Settings.WriteProtection.IsWriteProtected;

            // Log the results
            Console.WriteLine($"Requires password to open: {requiresOpenPassword}");
            Console.WriteLine($"Requires password to modify (write-protected): {requiresModifyPassword}");

            // Clean up
            loadedWorkbook.Dispose();
        }
    }
}