using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

class ParallelWorkbookEncryption
{
    // Derive a password from a seed string using SHA256 and return a hex representation
    private static string GetPassword(string seed)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(seed));
            // Use first 16 characters of the hex string as the password (adjust length as needed)
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            return sb.ToString().Substring(0, 16);
        }
    }

    static void Main()
    {
        // Define how many workbooks to create
        int workbookCount = 10;

        // Prepare a list of indices for parallel processing
        List<int> indices = new List<int>();
        for (int i = 1; i <= workbookCount; i++)
            indices.Add(i);

        // Encrypt each workbook in parallel, each with a distinct password
        Parallel.ForEach(indices, index =>
        {
            // Create a new workbook
            using (Workbook wb = new Workbook())
            {
                // Add simple data to identify the workbook
                Worksheet sheet = wb.Worksheets[0];
                sheet.Cells["A1"].PutValue($"Workbook #{index}");

                // Derive a unique password for this workbook
                string password = GetPassword($"WorkbookSeed_{index}");

                // Set the password (encryption) on the workbook settings
                wb.Settings.Password = password;

                // Save the encrypted workbook to disk
                string fileName = $"EncryptedWorkbook_{index}.xlsx";
                wb.Save(fileName);
            }
        });

        Console.WriteLine("All workbooks have been encrypted and saved.");
    }
}