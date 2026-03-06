using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class ValidateDigitalSignatureDemo
    {
        // Validates digital signatures in the specified XLSX file.
        public static void Run(string inputPath)
        {
            // Load the workbook from the given file path.
            Workbook workbook = new Workbook(inputPath);

            // Determine whether the workbook contains any digital signatures.
            bool isSigned = workbook.IsDigitallySigned;
            Console.WriteLine($"Workbook digitally signed: {isSigned}");

            if (isSigned)
            {
                // Retrieve the collection of digital signatures attached to the workbook.
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                // Iterate through each signature and display its details, including validity.
                foreach (DigitalSignature signature in signatures)
                {
                    Console.WriteLine($"Comment: {signature.Comments}");
                    Console.WriteLine($"Signed on: {signature.SignTime}");
                    Console.WriteLine($"Is valid: {signature.IsValid}");
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the XLSX file as a command‑line argument.");
                return;
            }

            string inputPath = args[0];
            ValidateDigitalSignatureDemo.Run(inputPath);
        }
    }
}