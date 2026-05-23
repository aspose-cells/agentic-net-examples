using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvSplitManifest
{
    class Program
    {
        static void Main()
        {
            // Paths for the source workbook, the intermediate CSV, split parts and manifest
            string sourceWorkbookPath = "source.xlsx";
            string intermediateCsvPath = "full.csv";
            string manifestPath = "manifest.txt";

            // Convert the entire workbook to a single CSV file using the provided ConversionUtility rule
            ConversionUtility.Convert(sourceWorkbookPath, intermediateCsvPath);

            // Define the maximum size for each split part (e.g., 1 MB)
            const int partSizeBytes = 1 * 1024 * 1024;

            // Read the whole CSV content into a byte array
            byte[] csvBytes = File.ReadAllBytes(intermediateCsvPath);

            // Calculate how many parts are needed
            int totalParts = (csvBytes.Length + partSizeBytes - 1) / partSizeBytes;

            // Prepare the manifest content
            using (StreamWriter manifestWriter = new StreamWriter(manifestPath))
            {
                for (int partIndex = 0; partIndex < totalParts; partIndex++)
                {
                    // Determine the slice of bytes for this part
                    int offset = partIndex * partSizeBytes;
                    int bytesToWrite = Math.Min(partSizeBytes, csvBytes.Length - offset);

                    // Build the part file name
                    string partFileName = $"part_{partIndex}.csv";

                    // Write the part to disk using standard FileStream (allowed I/O)
                    using (FileStream partStream = new FileStream(partFileName, FileMode.Create, FileAccess.Write))
                    {
                        partStream.Write(csvBytes, offset, bytesToWrite);
                    }

                    // Record the part name and its size in the manifest
                    FileInfo partInfo = new FileInfo(partFileName);
                    manifestWriter.WriteLine($"{partInfo.Name}\t{partInfo.Length} bytes");
                }
            }

            Console.WriteLine("CSV conversion, splitting, and manifest generation completed.");
        }
    }
}