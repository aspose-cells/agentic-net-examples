// Title: Check that a PDF exported from an Excel workbook using Aspose.Cells matches a specific byte size in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, saves it as a PDF into a MemoryStream, and validates that the stream length equals a predefined byte count. | Show how to compare the actual size of a PDF generated from a workbook with an expected size and output a success or mismatch message.
// Common Searches: C# Aspose.Cells verify PDF file size after saving workbook | How to assert exported PDF byte length matches expected value in .NET | Compare generated PDF size with predefined length using Aspose.Cells | Check PDF output size from Excel conversion Aspose.Cells C# example
// Tags: Aspose.Cells PDF export size validation | C# memory stream length check for PDF | compare expected and actual file size Aspose | assert exported document byte count .NET | verify workbook to PDF size Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook, saves it as a PDF into a MemoryStream using Aspose.Cells, then compares the stream's byte length to a predefined expected size and prints whether the sizes match.
class Program
{
    static void Main()
    {
        // Define the expected file size in bytes
        long expectedLength = 123456; // replace with the actual expected size

        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Export the workbook to a PDF file (as an example of a certificate export)
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Pdf);
            long actualLength = stream.Length;

            // Compare the exported file size with the expected length
            if (actualLength == expectedLength)
            {
                Console.WriteLine("Exported file size matches the expected length.");
            }
            else
            {
                Console.WriteLine($"File size mismatch. Expected: {expectedLength} bytes, Actual: {actualLength} bytes.");
            }
        }
    }
}
