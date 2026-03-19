using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the SpreadsheetML (Excel 2003 XML) file
            string filePath = "sample.xml";

            // Load the workbook from the XML file
            Workbook workbook = new Workbook(filePath);

            // Verify that the workbook contains a VBA project
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Check whether the VBA project is digitally signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If signed, verify the validity of the signature
                if (isSigned)
                {
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("VBA signature valid: " + isValid);
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the file.");
            }
        }
    }
}