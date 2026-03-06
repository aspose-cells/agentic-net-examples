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
            string workbookPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed and if the signature is valid
            bool isSigned = vbaProject.IsSigned;
            bool isValidSigned = vbaProject.IsValidSigned;

            // Prepare CSV content
            // Header: FilePath,IsSigned,IsValidSigned
            // Data row contains the results for the loaded workbook
            string csvContent = "FilePath,IsSigned,IsValidSigned" + Environment.NewLine;
            csvContent += $"\"{workbookPath}\",{isSigned},{isValidSigned}" + Environment.NewLine;

            // Output CSV file path
            string csvPath = "VbaSignatureReport.csv";

            // Write the CSV content to file
            File.WriteAllText(csvPath, csvContent);

            // Optional console output for immediate feedback
            Console.WriteLine($"VBA signature check completed. Results saved to '{csvPath}'.");
        }
    }
}