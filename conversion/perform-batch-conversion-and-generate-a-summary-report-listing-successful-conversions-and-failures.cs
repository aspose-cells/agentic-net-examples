using System;
using System.IO;
using System.Text;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class BatchConversion
    {
        // Performs batch conversion of Excel files to the specified format
        // and generates a plain‑text summary report.
        public static void Run(string[] sourceFiles, string targetExtension)
        {
            if (sourceFiles == null || sourceFiles.Length == 0)
            {
                Console.WriteLine("No source files provided.");
                return;
            }

            // Ensure the target extension starts with a dot (e.g., ".pdf")
            if (!targetExtension.StartsWith("."))
                targetExtension = "." + targetExtension;

            var report = new StringBuilder();
            report.AppendLine("Batch Conversion Report");
            report.AppendLine($"Run at: {DateTime.Now}");
            report.AppendLine(new string('=', 40));

            foreach (var sourcePath in sourceFiles)
            {
                // Validate source file existence
                if (!File.Exists(sourcePath))
                {
                    report.AppendLine($"{sourcePath} -> N/A : Source file not found");
                    continue;
                }

                // Build destination path by replacing the extension
                string destPath = Path.ChangeExtension(sourcePath, targetExtension);

                try
                {
                    // Convert using Aspose.Cells utility
                    ConversionUtility.Convert(sourcePath, destPath);

                    // Verify that the destination file was created
                    if (File.Exists(destPath))
                        report.AppendLine($"{sourcePath} -> {destPath} : Success");
                    else
                        report.AppendLine($"{sourcePath} -> {destPath} : Failed (output missing)");
                }
                catch (Exception ex)
                {
                    // Capture any conversion error
                    report.AppendLine($"{sourcePath} -> {destPath} : Failed ({ex.Message})");
                }
            }

            // Write the report to a text file in the same folder as the first source file
            string reportDir = Path.GetDirectoryName(sourceFiles[0]) ?? Directory.GetCurrentDirectory();
            string reportPath = Path.Combine(reportDir, "ConversionReport.txt");

            try
            {
                File.WriteAllText(reportPath, report.ToString());
                Console.WriteLine($"Conversion completed. Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to write report file: {ex.Message}");
            }
        }
    }

    class Program
    {
        // Entry point required for console application
        static void Main(string[] args)
        {
            try
            {
                // Example usage: convert all .xlsx files in a folder to .pdf
                string inputFolder = @"C:\InputFolder";

                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                string[] files = Directory.GetFiles(inputFolder, "*.xlsx");
                BatchConversion.Run(files, ".pdf");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}