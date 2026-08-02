// Title: Set deterministic TabId for worksheets using a name‑based hash in Aspose.Cells for .NET
// Description: Demonstrates how to compute a stable 32‑bit MD5 hash from each worksheet's name, assign it to the TabId property, and save the workbook. The approach guarantees identical TabId values for sheets with the same name across multiple workbooks.
// Keywords: Aspose.Cells TabId | deterministic worksheet identifier | hash worksheet name | MD5 to TabId | C# Excel TabId | consistent TabId across workbooks | .NET Aspose.Cells example
// Common Searches: Aspose.Cells set worksheet TabId by name | deterministic TabId for Excel sheets .NET | hash worksheet name for TabId Aspose.Cells | consistent TabId values across workbooks | C# compute integer hash for Excel tab
// Developer Intent: Generate a repeatable TabId for each worksheet based on its name.
// Use Cases: Synchronize sheets between separate workbooks by matching name‑derived TabId values. | Create a stable reference key for worksheets that survives file moves or format changes. | Enable fast lookup of specific tabs when loading large workbooks that use custom TabId identifiers.
// AI Prompts: Write C# code with Aspose.Cells that assigns ws.TabId = MD5Hash(ws.Name) for every worksheet. | Explain why a name‑based hash provides deterministic TabId values in Excel files. | Show an end‑to‑end example that creates a workbook, sets TabId using a 32‑bit hash, and saves it as XLSX.

using System;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Demonstrates how to compute a stable 32‑bit MD5 hash from each worksheet's name, assign it to the TabId property, and save the workbook. The approach guarantees identical TabId values for sheets with the same name across multiple workbooks.
    class Program
    {
        // Compute a deterministic 32‑bit hash from a string (worksheet name)
        static int ComputeNameHash(string name)
        {
            // Use MD5 to get a stable hash regardless of .NET version
            using (MD5 md5 = MD5.Create())
            {
                byte[] hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(name));
                // Take first 4 bytes and convert to Int32 (little‑endian)
                return BitConverter.ToInt32(hashBytes, 0);
            }
        }

        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Example: add a few worksheets with specific names
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Report");

            // Set TabId for each worksheet based on the hash of its name
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.TabId = ComputeNameHash(ws.Name);
                Console.WriteLine($"Worksheet \"{ws.Name}\" assigned TabId: {ws.TabId}");
            }

            // Save the workbook (choose any format you need)
            string outputPath = "TabIdDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
