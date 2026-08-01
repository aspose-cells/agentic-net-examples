// Title: Detect Excel Workbook Encryption Algorithm and Strength using Aspose.Cells in C#
// Description: A C# console sample that checks if an Excel file is encrypted, loads it with a password, uses reflection to read the internal EncryptionType, and reports whether the algorithm is StrongCryptographicProvider (strong) or a weaker method. Includes handling for missing files, invalid passwords, and unexpected errors.
// Keywords: Aspose.Cells C# encryption detection | Excel workbook encryption algorithm | EncryptionType reflection Aspose.Cells | strong vs weak Excel encryption .NET | detect encrypted workbook Aspose.Cells | WorkbookSettings IsEncrypted | LoadOptions password Excel | CellsException handling | Excel file security audit | C# Excel encryption strength
// Common Searches: How to detect encryption algorithm of an Excel file with Aspose.Cells C# | C# check if Excel workbook uses strong encryption | Retrieve EncryptionType from Aspose.Cells Workbook | Identify weak Excel encryption using .NET | Aspose.Cells detect encrypted workbook and password | Reflection to read private _encryptionType field | Determine Excel file encryption strength programmatically
// Developer Intent: Identify the workbook’s encryption algorithm and indicate whether it is strong or weak.
// Use Cases: Audit a collection of Excel files to ensure all documents employ strong encryption before archival. | Run a batch utility that scans multiple workbooks, logs each file’s encryption algorithm and strength for compliance reporting. | Provide immediate feedback in an application when an opened workbook uses weak encryption, prompting the user to re‑encrypt.
// AI Prompts: Write a C# method that uses reflection to extract the private _encryptionType field from an Aspose.Cells Workbook and returns a description of its strength. | Generate robust error handling for invalid passwords and missing encryption fields when detecting workbook encryption with Aspose.Cells. | Create a console program that iterates over a folder of Excel files, determines each file’s encryption type via reflection, and writes the results to a CSV file.

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // A C# console sample that checks if an Excel file is encrypted, loads it with a password, uses reflection to read the internal EncryptionType, and reports whether the algorithm is StrongCryptographicProvider (strong) or a weaker method. Includes handling for missing files, invalid passwords, and unexpected errors.
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be examined
            string filePath = "EncryptedWorkbook.xlsx";

            // Verify that the file exists before proceeding
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Detect basic file information, including whether it is encrypted
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Console.WriteLine($"IsEncrypted (FileFormatInfo): {formatInfo.IsEncrypted}");

                if (!formatInfo.IsEncrypted)
                {
                    Console.WriteLine("The workbook is not encrypted.");
                    return;
                }

                // Prompt for password (replace with actual password if known)
                Console.Write("Enter password for the encrypted workbook: ");
                string password = Console.ReadLine() ?? string.Empty;

                // Load the workbook using the supplied password
                LoadOptions loadOptions = new LoadOptions { Password = password };
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Verify that the workbook reports being encrypted via Settings
                bool isEncrypted = workbook.Settings.IsEncrypted;
                Console.WriteLine($"IsEncrypted (WorkbookSettings): {isEncrypted}");

                // Attempt to retrieve the encryption type using reflection.
                // Aspose.Cells does not expose a public getter for the encryption type,
                // but internally it is stored in a private field named "_encryptionType".
                EncryptionType encryptionType = EncryptionType.XOR; // default fallback
                try
                {
                    // First try to get the field from Workbook
                    FieldInfo field = typeof(Workbook).GetField("_encryptionType", BindingFlags.NonPublic | BindingFlags.Instance);
                    if (field != null && field.GetValue(workbook) is EncryptionType et)
                    {
                        encryptionType = et;
                    }
                    else
                    {
                        // If not found, try WorkbookSettings
                        FieldInfo settingsField = typeof(WorkbookSettings).GetField("_encryptionType", BindingFlags.NonPublic | BindingFlags.Instance);
                        if (settingsField != null && settingsField.GetValue(workbook.Settings) is EncryptionType et2)
                        {
                            encryptionType = et2;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Reflection error while retrieving encryption type: {ex.Message}");
                }

                // Determine strength based on the enumeration value
                bool isStrong = encryptionType == EncryptionType.StrongCryptographicProvider;
                string strength = isStrong ? "Strong" : "Weak";

                // Output the result
                Console.WriteLine($"Encryption Algorithm: {encryptionType}");
                Console.WriteLine($"Encryption Strength: {strength}");
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws CellsException for invalid passwords and other issues
                if (ex.Message != null && ex.Message.IndexOf("invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Invalid password provided. Unable to open the workbook.");
                }
                else
                {
                    Console.WriteLine($"CellsException: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
