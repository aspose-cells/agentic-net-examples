using System;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

class SetWorksheetTabIdDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Rename the default sheet and add additional sheets
        workbook.Worksheets[0].Name = "Summary";
        workbook.Worksheets.Add("Data");
        workbook.Worksheets.Add("Report");

        // Set TabId for each worksheet based on a deterministic hash of its name
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.TabId = ComputeDeterministicHash(ws.Name);
        }

        // Save the workbook
        workbook.Save("TabIdDemo.xlsx");
    }

    // Compute a deterministic 32‑bit integer hash from a string (MD5 first 4 bytes)
    static int ComputeDeterministicHash(string input)
    {
        using (MD5 md5 = MD5.Create())
        {
            byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));
            // Convert first 4 bytes to int (little‑endian)
            return BitConverter.ToInt32(hash, 0);
        }
    }
}