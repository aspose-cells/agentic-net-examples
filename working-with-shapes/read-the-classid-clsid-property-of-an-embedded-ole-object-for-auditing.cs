using System;
using System.IO;
using System.Runtime.InteropServices;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectClassIdAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains OLE objects
            string workbookPath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects in the current worksheet
                for (int i = 0; i < sheet.OleObjects.Count; i++)
                {
                    OleObject ole = sheet.OleObjects[i];

                    // Retrieve the ClassIdentifier (CLSID) as a byte array
                    byte[] clsidBytes = ole.ClassIdentifier;

                    // Convert the byte array to a readable GUID string if possible
                    string clsidString = ConvertClsIdToGuid(clsidBytes);

                    // Output audit information
                    Console.WriteLine($"Worksheet: {sheet.Name}");
                    Console.WriteLine($"OLE Object Index: {i}");
                    Console.WriteLine($"ProgID: {ole.ProgID}");
                    Console.WriteLine($"ClassIdentifier (CLSID): {clsidString}");
                    Console.WriteLine(new string('-', 50));
                }
            }
        }

        /// <summary>
        /// Converts a 16‑byte CLSID to a GUID string.
        /// If the byte array is not 16 bytes, returns a hex representation.
        /// </summary>
        private static string ConvertClsIdToGuid(byte[] clsidBytes)
        {
            if (clsidBytes == null || clsidBytes.Length == 0)
                return "None";

            if (clsidBytes.Length == 16)
            {
                // CLSID layout matches GUID layout; use Marshal to create a GUID
                Guid guid = new Guid(clsidBytes);
                return guid.ToString();
            }

            // Fallback: return hex string of the byte array
            return BitConverter.ToString(clsidBytes).Replace("-", "");
        }
    }
}