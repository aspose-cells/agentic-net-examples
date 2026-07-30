// Title: Detect Excel Workbook Encryption with Aspose.Cells for .NET (C#)
// Description: Load an Excel file using Aspose.Cells, read the Workbook.Settings.IsEncrypted flag, and output whether the file is password‑protected—all without supplying a password.
// Keywords: Aspose.Cells encryption detection | C# check Excel password protection | Workbook.Settings.IsEncrypted property | detect encrypted .xlsx with Aspose | Aspose.Cells .NET security | Excel encryption status C# | global Aspose.Cells encryption | US C# Excel encryption | India Aspose.Cells password check
// Common Searches: Aspose.Cells how to know if Excel file is encrypted | C# check if .xlsx is password protected using Aspose | IsWorkbookEncrypted Aspose.Cells example | Detect encrypted workbook without password .NET | Aspose.Cells encryption status code sample
// Developer Intent: The developer needs to load an existing Excel workbook and determine programmatically whether it is encrypted.
// Use Cases: Skip processing of password‑protected files in an automated import pipeline. | Prompt users for a password only when a workbook is identified as encrypted. | Record encryption flags for compliance auditing of uploaded spreadsheets.
// AI Prompts: Generate C# code with Aspose.Cells that reports the encryption state of an Excel file and asks for a password if needed. | Explain the behavior of Workbook.Settings.IsEncrypted when opening an encrypted workbook without a password, including any exceptions. | Show how to integrate encryption detection into a batch job that logs each file’s security status before further processing.

using System;
using Aspose.Cells;

// Load an Excel file using Aspose.Cells, read the Workbook.Settings.IsEncrypted flag, and output whether the file is password‑protected—all without supplying a password.
class Program
{
    static void Main()
    {
        // Path to the workbook file on disk
        string filePath = "sample.xlsx";

        // Load the workbook (no password required for just checking encryption)
        Workbook workbook = new Workbook(filePath);

        // Determine whether the workbook is encrypted
        bool isEncrypted = workbook.Settings.IsEncrypted;

        Console.WriteLine($"Workbook \"{filePath}\" is encrypted: {isEncrypted}");
    }
}
