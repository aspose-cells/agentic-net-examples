using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main()
        {
            const string filePath = "example.xlsm";

            // Ensure the workbook exists; create a minimal one if it does not.
            if (!File.Exists(filePath))
            {
                var tempWb = new Workbook();
                tempWb.Save(filePath, SaveFormat.Xlsm);
            }

            // Load the workbook that contains a VBA project
            var workbook = new Workbook(filePath);

            // Check whether the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + workbook.VbaProject.IsSigned);

            // If it is signed, verify the validity of the signature
            if (workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA Signature Valid: " + workbook.VbaProject.IsValidSigned);
            }

            // Reload the workbook from a memory stream to confirm the signature persists
            using (var stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsm);
                stream.Position = 0; // Reset stream position for reading

                var reloadedWorkbook = new Workbook(stream);
                Console.WriteLine("After reload - VBA Project Signed: " + reloadedWorkbook.VbaProject.IsSigned);
                Console.WriteLine("After reload - VBA Signature Valid: " + reloadedWorkbook.VbaProject.IsValidSigned);
            }
        }
    }
}