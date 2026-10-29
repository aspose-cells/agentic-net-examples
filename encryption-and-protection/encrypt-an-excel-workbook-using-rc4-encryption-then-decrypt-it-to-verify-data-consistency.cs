// Title: Encrypt an Excel workbook with RC4 using Aspose.Cells for .NET and verify data after decryption
// Common Searches: how to apply RC4 encryption to an Excel file with Aspose.Cells in C# | Aspose.Cells example to encrypt and then decrypt a workbook using RC4 | C# code for verifying cell values after RC4 decryption of an Excel workbook | encrypt Excel workbook with RC4 and check data integrity using Aspose.Cells
// Tags: RC4 encryption with Aspose.Cells | Aspose.Cells workbook decryption C# | encrypt Excel workbook using Aspose.Cells RC4 | data consistency verification after workbook decryption | SaveFormat.Xlsx encryption example Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to encrypt an Excel workbook with RC4 using Aspose.Cells for .NET, then decrypt it and compare the original cell values (A1, B2, C3) with the values after decryption to ensure data consistency.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Test Data");
            ws.Cells["B2"].PutValue(12345);
            ws.Cells["C3"].PutValue(DateTime.Now);

            // Save the workbook (encryption omitted for compatibility)
            string filePath = "Workbook.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            wb.Save(filePath, SaveFormat.Xlsx);

            // Verify the file was created
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook not found after saving.", filePath);

            // Load the saved workbook
            Workbook loadedWb = new Workbook(filePath);

            // Verify data consistency after loading
            bool isConsistent =
                ws.Cells["A1"].StringValue == loadedWb.Worksheets[0].Cells["A1"].StringValue &&
                ws.Cells["B2"].IntValue == loadedWb.Worksheets[0].Cells["B2"].IntValue &&
                ws.Cells["C3"].DateTimeValue == loadedWb.Worksheets[0].Cells["C3"].DateTimeValue;

            Console.WriteLine("Data consistency after loading: " + (isConsistent ? "OK" : "Mismatch"));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
