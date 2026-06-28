using System;
using Aspose.Cells;

namespace AsposeCellsDigitalSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing digitally signed workbook
            string sourcePath = "SignedWorkbook.xlsx";

            // Path where the copy will be saved
            string destinationPath = "SignedCopy.xlsx";

            // Load the signed workbook
            Workbook workbook = new Workbook(sourcePath);

            // Verify that the workbook is digitally signed
            if (workbook.IsDigitallySigned)
            {
                // Save the workbook to a new location.
                // The digital signature is preserved automatically.
                workbook.Save(destinationPath);
                Console.WriteLine($"Workbook saved to '{destinationPath}' with digital signature intact.");
            }
            else
            {
                Console.WriteLine("The source workbook is not digitally signed. No signature will be preserved.");
            }
        }
    }
}