// Title: C# Example: Validate Minimum Length of VBA Project Password Before Using Workbook.VbaProject.Protect with Aspose.Cells
// AI Prompts: Check the VBA project password length and raise an ArgumentException if it is shorter than the required minimum before calling workbook.VbaProject.Protect. | Add a configurable constant for the minimum password length and enforce it when protecting a VBA project in a C# Aspose.Cells workbook. | Implement a pre‑protect validation that ensures the password is not null or empty and meets the length rule, then protect the VBA project with the verified password.
// Common Searches: how to enforce a minimum password length when protecting a VBA project with Aspose.Cells in C# | C# Aspose.Cells validate VBA project password before calling Protect method | example of checking VBA password length prior to workbook.VbaProject.Protect | throw exception for short VBA password using Aspose.Cells API | configure minimum VBA password length in Aspose.Cells C# code
// Tags: VbaProject.Protect password validation C# | enforce password length rule Aspose.Cells | C# workbook VBA protection password check | Aspose.Cells argumentexception invalid VBA password | configure VBA password length Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The code loads an Excel workbook, verifies that the VBA project password meets a configurable minimum length, throws an ArgumentException if the rule is violated, protects the VBA project with the valid password using Workbook.VbaProject.Protect, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the VBA project password
            string vbaPassword = "mySecretPwd";

            // Minimum password length requirement
            const int minLength = 8;

            // Validate password length
            if (string.IsNullOrEmpty(vbaPassword) || vbaPassword.Length < minLength)
                throw new ArgumentException($"VBA project password must be at least {minLength} characters long.");

            // Protect the VBA project if it exists
            if (workbook.VbaProject != null)
            {
                // false = not read‑only; provide password to protect the project
                workbook.VbaProject.Protect(false, vbaPassword);
            }

            // Save the workbook with the protected VBA project
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
