using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionInfo
{
    public static class EncryptionHelper
    {
        /// <summary>
        /// Returns the name of the encryption algorithm used for the specified Excel file.
        /// For modern OOXML files (xlsx, xlsm, xlsb) Aspose.Cells uses AES encryption (StrongCryptographicProvider).
        /// For legacy binary files (xls) the exact algorithm cannot be determined via the public API,
        /// therefore the method returns a generic description.
        /// If the file is not encrypted, "None" is returned.
        /// </summary>
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="password">
        /// Optional password required to open the file.
        /// If the file is encrypted and a password is not supplied, the method will still detect encryption
        /// using the overload that does not require a password.
        /// </param>
        /// <returns>Encryption algorithm name or "None".</returns>
        public static string GetEncryptionAlgorithmName(string filePath, string password = null)
        {
            // Detect file format and encryption status.
            FileFormatInfo formatInfo;
            if (string.IsNullOrEmpty(password))
                formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            else
                formatInfo = FileFormatUtil.DetectFileFormat(filePath, password);

            // If the file is not encrypted, return "None".
            if (!formatInfo.IsEncrypted)
                return "None";

            // Determine algorithm based on file extension (OOXML vs legacy).
            string extension = System.IO.Path.GetExtension(filePath).ToLowerInvariant();

            // OOXML formats (xlsx, xlsm, xlsb) use AES encryption (StrongCryptographicProvider).
            if (extension == ".xlsx" || extension == ".xlsm" || extension == ".xlsb")
                return "AES (StrongCryptographicProvider)";

            // Legacy binary format (xls) uses one of the older algorithms.
            // The exact type cannot be retrieved via the public API, so we return a generic description.
            return "Legacy encryption (XOR/Compatible/EnhancedCryptographicProviderV1)";
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string filePath = "encrypted.xlsx";
            string password = "myPassword";

            string algorithm = EncryptionHelper.GetEncryptionAlgorithmName(filePath, password);
            Console.WriteLine($"Encryption algorithm for '{filePath}': {algorithm}");
        }
    }
}