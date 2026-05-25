using System;
using System.Collections;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue(123);
        cells["C1"].PutValue(DateTime.Now);
        cells["A2"].PutValue(3.14);
        cells["B2"].PutValue(true);

        // Concatenate string representations of all cells using Cells.GetEnumerator (rule)
        StringBuilder concatenated = new StringBuilder();
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell != null && cell.Value != null)
            {
                // Use DisplayStringValue for a formatted string representation
                concatenated.Append(cell.DisplayStringValue);
            }
        }

        // Compute a SHA256 checksum of the concatenated string
        byte[] hashBytes;
        using (SHA256 sha = SHA256.Create())
        {
            hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(concatenated.ToString()));
        }

        // Convert hash bytes to a hex string for display
        StringBuilder hex = new StringBuilder(hashBytes.Length * 2);
        foreach (byte b in hashBytes)
        {
            hex.AppendFormat("{0:x2}", b);
        }

        Console.WriteLine("Checksum (SHA256) of all cell values: " + hex.ToString());

        // Save the workbook (lifecycle: save)
        workbook.Save("ChecksumDemo.xlsx");
    }
}