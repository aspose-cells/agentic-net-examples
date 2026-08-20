// Title: Assign a deterministic TabId to worksheets by hashing their names – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds worksheets with custom names, generates a stable 32‑bit MD5 hash for each name, assigns the hash to the worksheet's TabId property, outputs the mapping, and saves the file.
// Keywords: Aspose.Cells TabId | deterministic worksheet identifier | hash worksheet name C# | stable TabId across workbooks | MD5 hash to int | C# .NET Excel automation
// Common Searches: set worksheet TabId based on name Aspose.Cells | generate consistent TabId for Excel sheets in C# | hash worksheet name to integer TabId | deterministic TabId for multiple workbooks | Aspose.Cells assign TabId programmatically
// Developer Intent: Create repeatable TabId values for worksheets by hashing their names.
// Use Cases: Synchronize worksheets across merged workbooks using identical TabId keys. | Implement fast lookup of sheets in UI navigation or API calls. | Maintain stable identifiers in version‑controlled Excel files.
// AI Prompts: Generate C# code that computes a 32‑bit MD5 hash of a worksheet name and sets ws.TabId in Aspose.Cells. | Explain how to guarantee the same TabId for a given worksheet name in different workbooks. | Show how to replace MD5 with SHA256 while still fitting the result into a 32‑bit TabId.

using System;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // This example creates a workbook, adds worksheets with custom names, generates a stable 32‑bit MD5 hash for each name, assigns the hash to the worksheet's TabId property, outputs the mapping, and saves the file.
    class Program
    {
        // Compute a deterministic 32‑bit hash from a string (e.g., worksheet name)
        static int GetDeterministicHash(string input)
        {
            // Use MD5 to get a stable hash across runs and platforms
            using (MD5 md5 = MD5.Create())
            {
                byte[] hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(input));
                // Take first 4 bytes and convert to int (little‑endian)
                return BitConverter.ToInt32(hashBytes, 0);
            }
        }

        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets with specific names
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Report");

            // Set TabId of each worksheet to a hash of its name
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.TabId = GetDeterministicHash(ws.Name);
                Console.WriteLine($"Worksheet \"{ws.Name}\" assigned TabId: {ws.TabId}");
            }

            // Save the workbook
            string outputPath = "DeterministicTabIdDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
