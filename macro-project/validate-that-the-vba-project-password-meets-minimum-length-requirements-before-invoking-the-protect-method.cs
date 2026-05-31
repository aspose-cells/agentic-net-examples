using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaProjectPasswordValidationDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Ensure a VBA project exists by adding a worksheet
        workbook.Worksheets.Add();

        // Define the password to protect the VBA project
        string password = "MySecurePwd123";

        // Minimum password length requirement
        const int MinPasswordLength = 8;

        // Validate password length before protecting
        if (string.IsNullOrEmpty(password) || password.Length < MinPasswordLength)
        {
            Console.WriteLine($"Password must be at least {MinPasswordLength} characters long.");
        }
        else
        {
            // Protect the VBA project (lock for viewing set to true)
            workbook.VbaProject.Protect(true, password);
            Console.WriteLine("VBA project protected successfully.");

            // Optional: verify that the password works using ValidatePassword
            bool isValid = workbook.VbaProject.ValidatePassword(password);
            Console.WriteLine($"Password validation result after protection: {isValid}");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}