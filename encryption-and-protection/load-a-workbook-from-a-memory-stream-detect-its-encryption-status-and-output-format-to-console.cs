// Title: Detect Excel Workbook Format and Encryption from a MemoryStream with Aspose.Cells for .NET
// Description: Shows how to create a Workbook, save it to a MemoryStream, and use Aspose.Cells.FileFormatUtil.DetectFileFormat to determine the file type (XLSX, XLS, CSV, etc.) and the encryption flag, then print the results to the console.
// Keywords: Aspose.Cells detect format | FileFormatUtil | memory stream encryption check | C# Excel encryption detection | detect workbook format .NET | Aspose.Cells IsEncrypted | read Excel from stream | detect file type without disk | Excel security check Aspose | in‑memory Excel format detection
// Common Searches: Aspose.Cells detect file format from MemoryStream | Check if Excel stream is encrypted using Aspose.Cells | FileFormatUtil DetectFileFormat C# example | How to read workbook format without saving to disk | Determine encryption status of XLSX in memory
// Developer Intent: Identify the format and encryption state of an Excel workbook held in a MemoryStream and display the information.
// Use Cases: Validate uploaded Excel files for encryption before server‑side processing. | Route in‑memory Excel data to appropriate handlers based on detected format. | Log encryption flags of generated workbooks for compliance auditing.
// AI Prompts: Provide C# code that reads a Stream and returns the file format type and IsEncrypted flag using Aspose.Cells. | Explain how to handle a workbook when DetectFileFormat reports IsEncrypted = true, including password opening. | Create a reusable method to detect Excel format and encryption from any Stream without writing to disk.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create a Workbook, save it to a MemoryStream, and use Aspose.Cells.FileFormatUtil.DetectFileFormat to determine the file type (XLSX, XLS, CSV, etc.) and the encryption flag, then print the results to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Optionally put some data into the workbook
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

        // Save the workbook to a memory stream in XLSX format
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0; // Reset stream position for reading

            // Detect file format and encryption status from the stream
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(stream);

            // Output detection results to the console
            Console.WriteLine($"Detected File Format Type: {fileInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {fileInfo.IsEncrypted}");
        }
    }
}
