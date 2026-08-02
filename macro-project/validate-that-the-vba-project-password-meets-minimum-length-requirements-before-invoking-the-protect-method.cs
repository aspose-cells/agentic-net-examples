using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Minimum password length requirement
            const int MinPasswordLength = 8;

            // Example password to protect the VBA project
            string password = "mySecretPwd";

            // Validate password length before protecting
            if (string.IsNullOrEmpty(password) || password.Length < MinPasswordLength)
            {
                throw new ArgumentException($"Password must be at least {MinPasswordLength} characters long.");
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Ensure a worksheet exists so that a VBA project is initialized
            workbook.Worksheets.Add();

            // Protect the VBA project (isLockedForViewing = false in this example)
            bool isLockedForViewing = false;
            workbook.VbaProject.Protect(isLockedForViewing, password);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
        }
    }
}