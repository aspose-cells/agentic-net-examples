using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookWriteProtectionCheck
    {
        public static void Run(string filePath, string passwordToValidate)
        {
            // Load the workbook without a password (only needed for write‑protection check)
            using (Workbook workbook = new Workbook(filePath))
            {
                // Access write‑protection settings
                WriteProtection writeProtection = workbook.Settings.WriteProtection;

                // Determine if the workbook is write‑protected (modification protected)
                bool isWriteProtected = writeProtection.IsWriteProtected;
                Console.WriteLine($"Workbook write‑protected: {isWriteProtected}");

                if (isWriteProtected)
                {
                    // Validate the supplied password against the write‑protection password
                    bool isPasswordValid = writeProtection.ValidatePassword(passwordToValidate);
                    Console.WriteLine($"Provided password valid: {isPasswordValid}");
                }
                else
                {
                    Console.WriteLine("Workbook is not write‑protected; password validation not required.");
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsExamples <filePath> <passwordToValidate>");
                return;
            }

            string filePath = args[0];
            string password = args[1];

            WorkbookWriteProtectionCheck.Run(filePath, password);
        }
    }
}