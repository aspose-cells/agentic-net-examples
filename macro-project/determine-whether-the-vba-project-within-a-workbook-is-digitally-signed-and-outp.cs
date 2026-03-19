using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel workbook (macro-enabled .xlsm file)
            string workbookPath = "input.xlsm";

            // Path to the output text file
            string outputTxtPath = "VbaSignatureStatus.txt";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is digitally signed
            bool isSigned = vbaProject != null && vbaProject.IsSigned;

            // Prepare the status message
            string statusMessage = isSigned
                ? "VBA project is signed."
                : "VBA project is NOT signed.";

            // Write the status to the text file
            File.WriteAllText(outputTxtPath, statusMessage);

            // Optional: also output to console for immediate feedback
            Console.WriteLine(statusMessage);
        }
    }
}