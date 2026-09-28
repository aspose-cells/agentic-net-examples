// Title: Determine if an Excel .xlsx workbook is digitally signed using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells and returns true when the workbook contains at least one digital signature. | Create a function that safely checks the Workbook.DigitalSignatureCollection and handles missing files or absent signature collection. | Provide a C# example that prints "Workbook signed: true/false" after loading an Excel file and evaluating its digital signature status with error handling.
// Common Searches: C# Aspose.Cells verify if an Excel file is signed | how to detect a signed workbook using Aspose.Cells .NET | read digital signatures from .xlsx with Aspose.Cells API | handle absent DigitalSignatureCollection in Aspose.Cells example
// Tags: Aspose.Cells access signature data | C# verify Excel workbook signature presence | load .xlsx and inspect workbook signatures with Aspose.Cells | handle missing Excel file errors Aspose.Cells | detect signature existence in Excel workbook .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook via Aspose.Cells, safely accesses its DigitalSignatureCollection to determine whether any digital signatures are present, and outputs the signed status while handling missing files and unavailable properties.
class Program
{
    static void Main()
    {
        const string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            // Use dynamic to access DigitalSignatureCollection without compile‑time binding
            dynamic workbook = new Workbook(filePath);

            bool isSigned = false;

            try
            {
                // Attempt to read the digital signature collection
                var signatureCollection = workbook.DigitalSignatureCollection;
                isSigned = signatureCollection != null && signatureCollection.Count > 0;
            }
            catch
            {
                // If the property is unavailable, treat the workbook as unsigned
                isSigned = false;
            }

            // Output the signature status
            Console.WriteLine($"Workbook signed: {isSigned}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
