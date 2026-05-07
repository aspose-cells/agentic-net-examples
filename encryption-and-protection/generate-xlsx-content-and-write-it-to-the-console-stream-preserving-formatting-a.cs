using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConsoleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data with basic formatting
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Apply bold style to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Save the workbook to a memory stream in XLSX format using the provided Save overload
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading

                // Write the binary XLSX content to the console output stream
                // This will output raw bytes; in a real scenario you might redirect to a file or further processing
                Stream consoleStream = Console.OpenStandardOutput();
                ms.CopyTo(consoleStream);
                consoleStream.Flush();
            }

            // Dispose workbook resources
            workbook.Dispose();
        }
    }
}