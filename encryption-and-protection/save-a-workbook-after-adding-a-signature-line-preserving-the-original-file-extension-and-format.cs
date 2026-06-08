using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook (replace with your actual file)
            string inputPath = "SampleWorkbook.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a signature line object
            SignatureLine signatureLine = new SignatureLine();

            // Add the signature line to the worksheet at row 5, column 1 (zero‑based index)
            // This uses the Shapes.AddSignatureLine method (no specific rule exists, so direct usage is allowed)
            worksheet.Shapes.AddSignatureLine(5, 1, signatureLine);

            // Determine the original file extension (including the dot)
            string originalExtension = Path.GetExtension(inputPath);

            // Convert the extension to the corresponding SaveFormat enum value
            SaveFormat saveFormat = FileFormatUtil.ExtensionToSaveFormat(originalExtension);

            // Prepare the output path – preserving the original extension
            string outputPath = "SignedCopy" + originalExtension;

            // Save the workbook using the Save(string, SaveFormat) overload (rule‑based)
            workbook.Save(outputPath, saveFormat);

            Console.WriteLine($"Workbook saved with signature to: {outputPath}");
        }
    }
}