using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureDemo
    {
        public static void Run()
        {
            // Load an Excel file in XLS format (Excel 97-2003) that may contain a VBA project
            Workbook workbook = new Workbook("sample.xls");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // If signed, verify whether the signature is valid
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValid);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed, so no signature validation is possible.");
            }

            // Optionally, save the workbook (preserving any existing signatures)
            workbook.Save("output.xls", SaveFormat.Excel97To2003);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaSignatureDemo.Run();
        }
    }
}