using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be inspected
            string sourcePath = "input.xlsx";

            // Detect file format and encryption status without opening the workbook
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath);
            Console.WriteLine($"Is file encrypted? {formatInfo.IsEncrypted}");

            // If the file is not encrypted, no further action is required
            if (!formatInfo.IsEncrypted)
            {
                Console.WriteLine("Workbook is not encrypted. No migration needed.");
                return;
            }

            // Load the encrypted workbook using the known password
            // (Replace "oldPassword" with the actual password for your file)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                Password = "oldPassword"
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Determine whether the workbook is using a deprecated encryption algorithm.
            // Aspose.Cells does not expose the encryption type directly after loading,
            // but the IsDefaultEncrypted flag indicates usage of the older default algorithm.
            bool usesDeprecatedAlgorithm = workbook.Settings.IsDefaultEncrypted;

            if (usesDeprecatedAlgorithm)
            {
                Console.WriteLine("Deprecated encryption algorithm detected. Migrating to StrongCryptographicProvider...");

                // Re‑encrypt the workbook with the recommended strong algorithm (AES 256)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
                workbook.Settings.Password = "newStrongPassword";

                // Save the migrated workbook
                string migratedPath = "migrated.xlsx";
                workbook.Save(migratedPath);
                Console.WriteLine($"Workbook re‑encrypted and saved to: {migratedPath}");
            }
            else
            {
                Console.WriteLine("Workbook uses a modern encryption algorithm. No migration required.");
            }

            // Clean up
            workbook.Dispose();
        }
    }
}