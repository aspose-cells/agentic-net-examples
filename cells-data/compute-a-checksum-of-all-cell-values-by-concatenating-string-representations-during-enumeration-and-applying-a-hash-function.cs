using System;
using System.Collections;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data (you can replace this with loading an existing file)
            cells["A1"].PutValue("Hello");
            cells["B1"].PutValue(123);
            cells["C1"].PutValue(DateTime.Now);
            cells["A2"].PutValue(3.14159);
            cells["B2"].PutValue(true);
            cells["C2"].PutValue("World");

            // Enumerate all cells, concatenate their string representations
            StringBuilder sb = new StringBuilder();
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                if (cell != null && cell.Value != null)
                {
                    // Use StringValue to get formatted string (or .Value.ToString())
                    sb.Append(cell.StringValue);
                }
            }

            // Compute SHA256 hash of the concatenated string
            byte[] hashBytes;
            using (SHA256 sha = SHA256.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(sb.ToString());
                hashBytes = sha.ComputeHash(inputBytes);
            }

            // Convert hash bytes to hexadecimal string
            StringBuilder hex = new StringBuilder(hashBytes.Length * 2);
            foreach (byte b in hashBytes)
                hex.AppendFormat("{0:x2}", b);

            Console.WriteLine("Checksum (SHA256): " + hex.ToString());

            // Optionally save the workbook
            workbook.Save("ChecksumDemo.xlsx");
        }
    }
}