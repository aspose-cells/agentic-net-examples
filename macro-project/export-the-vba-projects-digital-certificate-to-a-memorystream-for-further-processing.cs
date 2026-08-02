// Title: Export VBA Project Digital Certificate to a MemoryStream using Aspose.Cells for .NET
// Description: Loads an Excel workbook, checks if its VBA project is signed, extracts the certificate raw bytes via VbaProject.CertRawData, writes them to a MemoryStream, and returns the stream for further processing or saving.
// Keywords: Aspose.Cells | C# | .NET | VBA project certificate | CertRawData | MemoryStream export | signed VBA macro | extract digital certificate | Excel VBA security | global
// Common Searches: how to export VBA certificate to MemoryStream Aspose.Cells | retrieve signed VBA project certificate C# | Aspose.Cells VbaProject CertRawData example | save VBA digital certificate as binary file | check if VBA project is signed with Aspose
// Developer Intent: Obtain the digital certificate of a signed VBA project and deliver it as a MemoryStream for downstream use.
// Use Cases: Validate the extracted certificate against a trusted store before macro execution. | Send the MemoryStream to a web service for remote verification. | Archive the certificate bytes to a .bin file for compliance auditing.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect a signed VBA project and return its certificate as a MemoryStream. | Create a method that extracts VbaProject.CertRawData, writes it to a .bin file, and includes robust error handling. | Show how to post the MemoryStream containing a VBA certificate to a REST API with HttpClient.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, checks if its VBA project is signed, extracts the certificate raw bytes via VbaProject.CertRawData, writes them to a MemoryStream, and returns the stream for further processing or saving.
    public class ExportVbaCertificateToMemoryStream
    {
        /// <param name="workbookPath">Path to the Excel file that contains a signed VBA project.</param>
        /// <returns>A MemoryStream containing the certificate raw data, or null if the project is not signed.</returns>
        public static MemoryStream Run(string workbookPath)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return null;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed
                if (!vbaProject.IsSigned)
                {
                    Console.WriteLine("The VBA project is not signed. No certificate data available.");
                    return null;
                }

                // Get the certificate raw data (byte array)
                byte[] certData = vbaProject.CertRawData;

                if (certData == null || certData.Length == 0)
                {
                    Console.WriteLine("Certificate raw data is empty.");
                    return null;
                }

                // Write the certificate data into a MemoryStream
                MemoryStream certStream = new MemoryStream();
                certStream.Write(certData, 0, certData.Length);
                certStream.Position = 0; // Reset position for downstream consumers

                Console.WriteLine($"Certificate exported to MemoryStream. Length: {certStream.Length} bytes.");

                return certStream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
                return null;
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage: provide the path to the Excel file as the first argument
            string workbookPath = args.Length > 0 ? args[0] : "SignedVbaProject.xlsx";

            Console.WriteLine($"Processing workbook: {workbookPath}");

            MemoryStream result = ExportVbaCertificateToMemoryStream.Run(workbookPath);

            if (result != null)
            {
                // Optionally, save the certificate data to a file for verification
                string outputPath = "VbaCertificate.bin";
                try
                {
                    File.WriteAllBytes(outputPath, result.ToArray());
                    Console.WriteLine($"Certificate data saved to {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write certificate file: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No certificate data was exported.");
            }
        }
    }
}
