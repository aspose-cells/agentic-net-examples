using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath;
            if (args.Length > 0)
                filePath = args[0];
            else
            {
                Console.Write("Enter workbook path: ");
                filePath = Console.ReadLine() ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("No workbook path provided.");
                return;
            }

            SignatureValidator.Validate(filePath);
        }
    }

    public static class SignatureValidator
    {
        public static void Validate(string filePath)
        {
            using (Workbook workbook = new Workbook(filePath))
            {
                if (!workbook.IsDigitallySigned)
                {
                    Console.WriteLine("Workbook is not digitally signed.");
                    return;
                }

                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                if (signatures == null)
                {
                    Console.WriteLine("No digital signatures collection found.");
                    return;
                }

                bool anySignature = false;
                bool allValid = true;
                int index = 0;

                foreach (DigitalSignature sig in signatures)
                {
                    anySignature = true;
                    index++;
                    Console.WriteLine($"Signature #{index}:");
                    Console.WriteLine($"  Comments : {sig.Comments}");
                    Console.WriteLine($"  Sign Time: {sig.SignTime}");
                    Console.WriteLine($"  Is Valid : {sig.IsValid}");

                    if (!sig.IsValid)
                        allValid = false;
                }

                if (!anySignature)
                {
                    Console.WriteLine("No digital signatures found.");
                    return;
                }

                if (allValid)
                {
                    Console.WriteLine("All digital signatures are valid. Proceeding with further processing.");
                    // TODO: Insert workbook processing logic here
                }
                else
                {
                    Console.WriteLine("One or more digital signatures are invalid. Processing aborted.");
                }
            }
        }
    }
}