using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureCsvDemo
    {
        public static void Run(string workbookPath, string csvPath)
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed and if the signature is valid
            bool isSigned = vbaProject.IsSigned;
            bool isValidSigned = vbaProject.IsValidSigned;

            // Prepare CSV line
            string csvLine = $"{Path.GetFileName(workbookPath)},{isSigned},{isValidSigned}";

            // Write header if CSV does not exist, then append the result
            bool writeHeader = !File.Exists(csvPath);
            using (StreamWriter writer = new StreamWriter(csvPath, true))
            {
                if (writeHeader)
                {
                    writer.WriteLine("Workbook,IsSigned,IsValidSigned");
                }
                writer.WriteLine(csvLine);
            }

            // Save the workbook to a memory stream to demonstrate the save rule
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                // Reload to verify the signature persists (optional verification)
                Workbook verifyWorkbook = new Workbook(ms);
                // No further action needed; verification can be done via debug or logs
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // args[0] - path to the input workbook (XLSM)
            // args[1] - path to the output CSV file
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <workbookPath> <csvPath>");
                return;
            }

            string workbookPath = args[0];
            string csvPath = args[1];

            VbaSignatureCsvDemo.Run(workbookPath, csvPath);
        }
    }
}