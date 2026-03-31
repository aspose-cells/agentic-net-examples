using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureVerification
{
    public class VbaSignatureVerifier
    {
        /// <summary>
        /// Verifies the VBA project signature of the specified Excel file and returns the result as a JSON string.
        /// </summary>
        /// <param name="excelFilePath">Full path to the Excel workbook (xlsm, xlsb, etc.).</param>
        /// <returns>JSON string containing signature verification information.</returns>
        public static string VerifyVbaSignature(string excelFilePath)
        {
            // Load the workbook (load rule)
            Workbook workbook = new Workbook(excelFilePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Prepare result object
            var result = new
            {
                IsSigned = vbaProject.IsSigned,
                IsValidSigned = vbaProject.IsValidSigned,
                // If signed, include certificate raw data as Base64 string; otherwise null
                CertificateBase64 = vbaProject.IsSigned && vbaProject.CertRawData != null
                                    ? Convert.ToBase64String(vbaProject.CertRawData)
                                    : null
            };

            // Serialize result to JSON
            string jsonResult = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });

            return jsonResult;
        }

        // Example usage
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file as a command‑line argument.");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                string verificationJson = VerifyVbaSignature(filePath);
                Console.WriteLine("VBA Signature Verification Result:");
                Console.WriteLine(verificationJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during verification: {ex.Message}");
            }
        }
    }
}