using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsInMemoryDemo
{
    public class WorkbookSerializer
    {
        /// <summary>
        /// Creates a sample workbook, saves it to a memory stream in XLSX format,
        /// and returns the resulting byte array.
        /// </summary>
        /// <returns>Byte array containing the XLSX file.</returns>
        public static byte[] GetWorkbookAsXlsxBytes()
        {
            // Initialize a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(456);

            // Save the workbook to a MemoryStream using the Save(Stream, SaveFormat) overload
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                // Ensure the stream position is at the beginning before reading
                ms.Position = 0;
                // Return the underlying byte array
                return ms.ToArray();
            }
        }

        // Example usage
        public static void Main()
        {
            byte[] xlsxBytes = GetWorkbookAsXlsxBytes();
            Console.WriteLine($"Workbook serialized to memory. Byte array length: {xlsxBytes.Length}");
            // The byte array can now be transmitted, stored, or processed further without touching the file system.
        }
    }
}