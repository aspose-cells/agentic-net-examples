// Title: C# – Validate VBA Project Password Length Before Using Aspose.Cells Protect
// Description: Demonstrates how to create a macro‑enabled workbook, enforce a configurable minimum password length, and call Workbook.VbaProject.Protect only when the password meets the requirement, then save the file as .xlsm.
// Keywords: Aspose.Cells | C# VBA project protection | password length validation | minimum password requirement | macro‑enabled workbook | Workbook.VbaProject.Protect | XLSM file generation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells validate VBA password length | C# protect VBA project only if password is long enough | how to enforce minimum password for VBA project in Aspose.Cells | protect macro workbook with password validation .NET | sample code for Workbook.VbaProject.Protect password check
// Developer Intent: Check that a VBA project password satisfies a minimum length before invoking the Protect method in Aspose.Cells.
// Use Cases: Enforcing security policy by rejecting short passwords when protecting a VBA project. | Automating workbook creation where the VBA project is locked only after successful password validation. | Providing user feedback and aborting protection when the password does not meet length criteria.
// AI Prompts: Write C# code using Aspose.Cells that verifies a VBA project password meets a configurable minimum length before calling Protect. | Show how to handle an invalid‑password scenario when protecting a VBA project in a macro‑enabled workbook with Aspose.Cells. | Generate a logging example that records a warning and skips VBA project protection if the password is too short.

using System;
using Aspose.Cells;

// Demonstrates how to create a macro‑enabled workbook, enforce a configurable minimum password length, and call Workbook.VbaProject.Protect only when the password meets the requirement, then save the file as .xlsm.
class VbaProjectPasswordValidationDemo
{
    static void Main()
    {
        // Create a new workbook (this also creates a VBA project)
        Workbook workbook = new Workbook();
        // Ensure there is at least one worksheet so the VBA project is initialized
        workbook.Worksheets.Add();

        // Define the password to protect the VBA project
        string password = "MySecurePwd123";
        // Minimum required length for the password
        int minLength = 8;

        // Validate password length before calling Protect
        if (string.IsNullOrEmpty(password) || password.Length < minLength)
        {
            Console.WriteLine($"Password must be at least {minLength} characters long.");
            return; // Abort protection if the password does not meet the requirement
        }

        // Protect the VBA project and lock it for viewing (set to true as an example)
        bool lockForViewing = true;
        workbook.VbaProject.Protect(lockForViewing, password);

        // Save the workbook as a macro-enabled file
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
        Console.WriteLine("VBA project protected successfully.");
    }
}
