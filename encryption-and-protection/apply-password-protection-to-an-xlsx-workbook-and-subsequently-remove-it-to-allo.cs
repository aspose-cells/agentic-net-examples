using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Step 2: Protect the workbook with a password (protecting structure and windows)
            string password = "Secret123";
            workbook.Protect(ProtectionType.All, password);

            // Step 3: Save the password‑protected workbook
            string protectedPath = "protected_workbook.xlsx";
            workbook.Save(protectedPath, SaveFormat.Xlsx);
            workbook.Dispose();

            // Step 4: Load the protected workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook protectedWb = new Workbook(protectedPath, loadOptions);

            // Step 5: Unprotect the workbook using the same password
            protectedWb.Unprotect(password);

            // Step 6: Save the now unprotected workbook
            string unprotectedPath = "unprotected_workbook.xlsx";
            protectedWb.Save(unprotectedPath, SaveFormat.Xlsx);
            protectedWb.Dispose();

            Console.WriteLine("Workbook protected and then unprotected successfully.");
        }
    }
}