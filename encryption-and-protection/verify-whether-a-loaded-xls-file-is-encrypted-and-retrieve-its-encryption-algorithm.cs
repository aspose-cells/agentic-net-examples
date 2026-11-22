// Title: Determine if an XLS workbook is encrypted and retrieve its encryption algorithm with Aspose.Cells for .NET (C#)
// Common Searches: how to verify encryption of an XLS file using Aspose.Cells in C# | retrieve the encryption algorithm of a protected Excel workbook with Aspose.Cells .NET | check Workbook.IsEncrypted property in Aspose.Cells without a password | use reflection to get WorkbookSettings EncryptionAlgorithm in C# | detect encrypted XLS and read its encryption method with Aspose.Cells API
// Tags: Aspose.Cells detect encrypted XLS workbook | Workbook.IsEncrypted property C# | WorkbookSettings EncryptionAlgorithm extraction | load XLS file without password Aspose.Cells | identify Excel encryption type .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLS file using Aspose.Cells for .NET, checks the Workbook.IsEncrypted flag (via reflection for version safety), and if encrypted extracts the EncryptionAlgorithm from WorkbookSettings, then prints the encryption status and algorithm while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Path to the XLS file to be examined
        string filePath = "sample.xls";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook (no password supplied)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(filePath, loadOptions);

            bool isEncrypted = false;
            string algorithmInfo = "N/A";

            // Try to obtain the IsEncrypted property via reflection (covers different API versions)
            var isEncryptedProp = typeof(Workbook).GetProperty("IsEncrypted");
            if (isEncryptedProp != null)
            {
                isEncrypted = (bool)isEncryptedProp.GetValue(workbook);
            }

            // If the workbook is encrypted, attempt to read the encryption algorithm via reflection
            if (isEncrypted)
            {
                var encryptionAlgProp = typeof(WorkbookSettings).GetProperty("EncryptionAlgorithm");
                if (encryptionAlgProp != null)
                {
                    var algValue = encryptionAlgProp.GetValue(workbook.Settings);
                    algorithmInfo = algValue?.ToString() ?? "Unknown";
                }
            }

            Console.WriteLine($"Is Encrypted: {isEncrypted}");
            if (isEncrypted)
            {
                Console.WriteLine($"Encryption Algorithm: {algorithmInfo}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
