// Title: Check if an Excel workbook is digitally signed with Aspose.Cells for .NET
// Description: Loads an Excel file using Aspose.Cells, confirms the file exists, reads the Workbook.IsDigitallySigned property, and outputs the signature status while handling exceptions.
// Keywords: Aspose.Cells | C# digital signature | Workbook.IsDigitallySigned | Excel signature detection | verify Excel workbook signature | Aspose.Cells example | digital signature status | Excel file signed check
// Common Searches: Aspose.Cells how to check if Excel file is signed | C# read digital signature of workbook | IsDigitallySigned property example | detect signed Excel workbook .NET | verify Excel digital signature programmatically
// Developer Intent: Determine whether a given Excel workbook contains a digital signature.
// Use Cases: Validate incoming Excel reports are signed before automated processing. | Flag or skip unsigned workbooks during batch imports to enforce trust policies. | Record signature status of each opened workbook for audit and compliance.
// AI Prompts: Generate C# code using Aspose.Cells that checks Workbook.IsDigitallySigned and reports missing file errors. | Show how to extract signer name and signing timestamp from a digitally signed workbook with Aspose.Cells. | Create a reusable method that returns the signature flag and throws a custom exception when the workbook is unsigned.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSignatureCheck
{
    // Loads an Excel file using Aspose.Cells, confirms the file exists, reads the Workbook.IsDigitallySigned property, and outputs the signature status while handling exceptions.
    public class Program
    {
        public static void Main()
        {
            // Path to the Excel file to be examined
            string filePath = "Test_DigiSign.xlsx";

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Retrieve the digital signature status
                bool isSigned = workbook.IsDigitallySigned;

                // Display the result
                Console.WriteLine($"Workbook is digitally signed: {isSigned}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
