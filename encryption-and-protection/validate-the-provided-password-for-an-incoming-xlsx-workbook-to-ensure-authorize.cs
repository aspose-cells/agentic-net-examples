using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    public class Program
    {
        public static void Main()
        {
            // Path to the incoming XLSX workbook
            string filePath = "incoming.xlsx";

            // Password supplied by the user (could be obtained from UI, args, etc.)
            string suppliedPassword = "UserProvidedPassword";

            // Ensure the workbook exists; if not, create a sample encrypted workbook
            if (!File.Exists(filePath))
            {
                Workbook sampleWb = new Workbook();
                sampleWb.Worksheets[0].Name = "SampleSheet";
                sampleWb.Settings.Password = suppliedPassword; // encrypt with the same password
                sampleWb.Save(filePath);
            }

            // First, verify whether the supplied password can decrypt the workbook
            bool isPasswordCorrect;
            using (FileStream stream = File.OpenRead(filePath))
            {
                // FileFormatUtil.VerifyPassword checks the password for encrypted OOXML files
                isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, suppliedPassword);
            }

            Console.WriteLine($"Password validation result: {isPasswordCorrect}");

            if (isPasswordCorrect)
            {
                // Load the workbook using the correct password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = suppliedPassword
                };

                // Load the workbook
                Workbook workbook = new Workbook(filePath, loadOptions);

                // At this point the workbook is successfully decrypted and can be used
                Console.WriteLine("Workbook loaded successfully. First sheet name: " +
                                  workbook.Worksheets[0].Name);

                // Example: Save a copy without password (optional)
                workbook.Settings.Password = null; // remove encryption
                workbook.Save("decrypted_copy.xlsx");
            }
            else
            {
                Console.WriteLine("Unable to open the workbook. Incorrect password.");
            }
        }
    }
}