// Title: Parallel encryption of multiple Excel workbooks with SHA‑256 derived passwords using Aspose.Cells for .NET
// Description: C# sample that creates several workbooks, generates a unique password for each from the first 8 bytes of a SHA‑256 hash, applies StrongCryptographicProvider (128‑bit) encryption via Workbook.Settings.Password, and saves the files concurrently with Parallel.ForEach.
// Keywords: Aspose.Cells | C# | .NET | Excel encryption | parallel workbook processing | SHA256 password generation | StrongCryptographicProvider | 128‑bit encryption | Workbook.Settings.Password | Parallel.ForEach | batch Excel protection | GitHub example
// Common Searches: encrypt multiple Excel files in parallel C# Aspose.Cells | generate unique password for each workbook using SHA256 | set strong encryption options when saving Excel with Aspose.Cells | parallel batch workbook protection .NET | example code for Aspose.Cells workbook encryption GitHub
// Developer Intent: Secure a collection of Excel workbooks at the same time, assigning each a distinct hash‑derived password.
// Use Cases: Batch‑process financial reports, giving each file a password based on its dataset identifier. | Run a multi‑threaded service that creates and protects temporary spreadsheets for different users. | Automate export of confidential data sets where each export must have a unique strong password.
// AI Prompts: Generate C# code that uses Aspose.Cells to encrypt a list of Excel files in parallel, deriving each password from the first 8 bytes of a SHA‑256 hash. | Show how to configure Aspose.Cells workbook encryption with StrongCryptographicProvider and a 128‑bit key, then save the workbook using Workbook.Settings.Password. | Explain best practices for parallelizing workbook creation and encryption with Parallel.ForEach while handling exceptions in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

// C# sample that creates several workbooks, generates a unique password for each from the first 8 bytes of a SHA‑256 hash, applies StrongCryptographicProvider (128‑bit) encryption via Workbook.Settings.Password, and saves the files concurrently with Parallel.ForEach.
class Program
{
    static void Main()
    {
        // Define workbook file paths and a base string for each password derivation
        var workbookInfos = new List<(string FilePath, string BaseString)>
        {
            ("Workbook1.xlsx", "DataSetA"),
            ("Workbook2.xlsx", "DataSetB"),
            ("Workbook3.xlsx", "DataSetC")
        };

        // Encrypt each workbook in parallel
        Parallel.ForEach(workbookInfos, info =>
        {
            // Derive a password from a SHA256 hash of the base string
            string password = DerivePassword(info.BaseString);

            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // Add sample data to the first worksheet
            wb.Worksheets[0].Cells["A1"].PutValue($"Sample data for {info.BaseString}");

            // Set the workbook encryption password (lifecycle rule: set property)
            wb.Settings.Password = password;

            // Optionally specify strong encryption options
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook (lifecycle rule: save)
            wb.Save(info.FilePath);
        });

        Console.WriteLine("All workbooks have been encrypted.");
    }

    // Helper method to create a password from the first 8 bytes of a SHA256 hash
    static string DerivePassword(string input)
    {
        using (SHA256 sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(input));
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 8; i++) // use 8 bytes => 16 hex characters
            {
                sb.Append(hash[i].ToString("x2"));
            }
            return sb.ToString();
        }
    }
}
