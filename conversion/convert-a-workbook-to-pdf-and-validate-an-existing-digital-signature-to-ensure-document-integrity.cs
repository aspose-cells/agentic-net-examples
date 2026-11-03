// Title: C# example: Convert an Excel workbook to PDF and validate all digital signatures with Aspose.Cells
// Common Searches: how to convert an .xlsx file to PDF and check digital signatures using Aspose.Cells in C# | Aspose.Cells C# validate all digital signatures before exporting workbook to PDF | C# code to load Excel, verify digital signature collection, and save as PDF with Aspose.Cells | using reflection to access DigitalSignatureCollection ValidateAll in Aspose.Cells .NET | sample program to ensure Excel workbook integrity and PDF conversion with Aspose.Cells
// Tags: Aspose.Cells workbook to PDF conversion C# | Aspose.Cells digital signature validation .NET | ValidateAll method DigitalSignatureCollection Aspose.Cells | C# Excel to PDF export with signature check | reflection access Aspose.Cells DigitalSignatureCollection

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an Excel file into an Aspose.Cells Workbook, uses reflection to call the DigitalSignatureCollection.ValidateAll method to verify all signatures, prints the validation result, and then saves the workbook as a PDF, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Validate all digital signatures if the API is available
            bool signaturesAreValid = false;
            try
            {
                PropertyInfo dsProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                if (dsProp != null)
                {
                    object dsCollection = dsProp.GetValue(workbook);
                    MethodInfo validateMethod = dsCollection?.GetType().GetMethod("ValidateAll");
                    if (validateMethod != null)
                    {
                        signaturesAreValid = (bool)validateMethod.Invoke(dsCollection, null);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Digital signature validation error: {ex.Message}");
            }

            Console.WriteLine($"Digital signatures valid: {signaturesAreValid}");

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
