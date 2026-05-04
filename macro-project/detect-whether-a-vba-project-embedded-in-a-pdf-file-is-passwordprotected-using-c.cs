using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaPdfDetection
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file (extracted from PDF by any external means)
            string excelPath = "sample.xlsx";

            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(excelPath)}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project exists and whether it is protected with a password
            bool isVbaProtected = vbaProject != null && vbaProject.IsProtected;

            Console.WriteLine($"VBA project password‑protected: {isVbaProtected}");

            // Optional: if you need to verify a specific password, use ValidatePassword
            // string passwordToTest = "myPassword";
            // bool isPasswordCorrect = vbaProject?.ValidatePassword(passwordToTest) ?? false;
            // Console.WriteLine($"Password '{passwordToTest}' is valid: {isPasswordCorrect}");
        }
    }
}