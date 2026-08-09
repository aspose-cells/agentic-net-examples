// Title: Compute SHA‑256 Hash of an Encrypted Aspose.Cells Workbook in C#
// Description: Shows how to create an Aspose.Cells workbook, protect it with a password, save it to a MemoryStream, calculate a SHA‑256 checksum of the encrypted bytes using .NET's SHA256 class, and output the hex hash together with the IsEncrypted flag for integrity verification.
// Keywords: Aspose.Cells | C# SHA-256 | encrypted workbook checksum | Excel file integrity | password protected Excel hash | SHA256 Aspose.Cells | verify workbook integrity | compute workbook hash .NET
// Common Searches: C# compute SHA256 of password protected Excel | Aspose.Cells hash encrypted workbook | verify integrity of encrypted Excel file using .NET | how to get SHA256 checksum after workbook encryption | Aspose.Cells IsEncrypted flag example
// Developer Intent: Obtain a SHA‑256 checksum of a password‑protected workbook to confirm its integrity after saving.
// Use Cases: Store the hash alongside the encrypted file for later tamper detection. | Log the checksum in an audit trail when generating confidential reports. | Compare the computed hash with a known value before transmitting the workbook over a network. | Automate integrity checks in CI pipelines for generated Excel assets.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, applies a password, saves it to a stream, and returns the SHA‑256 hash as a lowercase hex string. | Show how to compare a newly generated SHA‑256 hash of an encrypted Excel file with a previously saved hash to detect modifications. | Explain how to use workbook.Settings.IsEncrypted together with a SHA‑256 checksum to build an integrity‑verification routine.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// Shows how to create an Aspose.Cells workbook, protect it with a password, save it to a MemoryStream, calculate a SHA‑256 checksum of the encrypted bytes using .NET's SHA256 class, and output the hex hash together with the IsEncrypted flag for integrity verification.
class ComputeWorkbookHash
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

        // Encrypt the workbook with a password
        workbook.Settings.Password = "mySecretPassword";

        // Save the encrypted workbook to a memory stream
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);

            // Compute SHA‑256 hash of the encrypted file bytes
            byte[] fileBytes = stream.ToArray();
            byte[] hashBytes;
            using (SHA256 sha256 = SHA256.Create())
            {
                hashBytes = sha256.ComputeHash(fileBytes);
            }

            // Convert hash to a hexadecimal string for display
            string hashHex = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            Console.WriteLine("SHA‑256 hash of encrypted workbook: " + hashHex);
        }

        // Verify encryption status (optional)
        Console.WriteLine("Workbook IsEncrypted: " + workbook.Settings.IsEncrypted);
    }
}
