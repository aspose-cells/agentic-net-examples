// Title: Detect Excel Workbook Encryption Algorithm and Strength with Aspose.Cells (C#)
// Description: C# example that loads an Excel file using Aspose.Cells, checks if the workbook is encrypted, retrieves the internal EncryptionType via reflection, maps the enum to strong or weak categories, and logs the result. Includes a fallback message when the property cannot be accessed.
// Keywords: Aspose.Cells encryption detection | Excel workbook EncryptionType | C# identify Excel encryption algorithm | strong vs weak Excel encryption .NET | reflection read internal property Aspose.Cells | Workbook.Settings.IsEncrypted | EncryptionType enum Aspose.Cells | Excel file security assessment
// Common Searches: How to get encryption algorithm of an Excel file using Aspose.Cells | Determine if Excel workbook uses strong encryption in C# | Read EncryptionType property with reflection Aspose.Cells | Check workbook encryption strength without password | Aspose.Cells detect weak encryption
// Developer Intent: Identify the encryption algorithm of an Excel workbook and report whether it is classified as strong or weak.
// Use Cases: Compliance auditing: automatically flag workbooks that use weak encryption. | Pre‑processing validation: decide whether to decrypt or reject files based on encryption strength. | Logging and monitoring: record encryption details for security dashboards. | Fallback handling: provide a default message when the EncryptionType property is unavailable.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, uses reflection to obtain the EncryptionType enum, and prints "strong" or "weak" based on the value. | Explain how Aspose.Cells maps EncryptionType values to cryptographic strength categories. | Suggest a non‑reflection approach to evaluate workbook encryption strength in Aspose.Cells.

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// C# example that loads an Excel file using Aspose.Cells, checks if the workbook is encrypted, retrieves the internal EncryptionType via reflection, maps the enum to strong or weak categories, and logs the result. Includes a fallback message when the property cannot be accessed.
class IdentifyEncryption
{
    static void Main()
    {
        // Path to the workbook to analyze
        string filePath = "sample.xlsx";

        // Load the workbook (no password needed for detection)
        Workbook workbook = new Workbook(filePath);

        // Check if the workbook is encrypted
        bool isEncrypted = workbook.Settings.IsEncrypted;
        Console.WriteLine($"Workbook encrypted: {isEncrypted}");

        if (isEncrypted)
        {
            // Try to obtain the encryption type via reflection.
            // Aspose.Cells may expose an internal property named "EncryptionType" in WorkbookSettings.
            PropertyInfo encProp = workbook.Settings.GetType()
                .GetProperty("EncryptionType", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);

            if (encProp != null)
            {
                // Cast the retrieved value to the public EncryptionType enum.
                EncryptionType encType = (EncryptionType)encProp.GetValue(workbook.Settings);
                Console.WriteLine($"Encryption algorithm: {encType}");

                // Determine strength based on the enum value.
                bool isStrong = encType == EncryptionType.StrongCryptographicProvider ||
                                encType == EncryptionType.EnhancedCryptographicProviderV1;

                Console.WriteLine(isStrong ? "Encryption is strong." : "Encryption is weak.");
            }
            else
            {
                // Fallback when the property is not accessible.
                Console.WriteLine("Unable to determine the exact encryption algorithm. Assuming default strong encryption for modern formats.");
            }
        }
        else
        {
            Console.WriteLine("Workbook is not encrypted.");
        }
    }
}
