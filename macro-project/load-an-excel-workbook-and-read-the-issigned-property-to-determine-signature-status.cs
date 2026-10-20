// Title: How to detect a digital signature in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check if .xlsx file has a digital signature | read digital signature count from workbook using Aspose.Cells .NET | use reflection to access DigitalSignatureCollection in Aspose.Cells | determine signed status of an Excel file programmatically with Aspose.Cells | verify workbook signature existence in C# Aspose.Cells example
// Tags: Aspose.Cells signature collection size | C# workbook signed status detection | reflection access DigitalSignatureCollection | load Excel file and verify signature | check IsSigned property Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The sample loads an Excel workbook from a specified path, safely checks for file existence, and uses reflection to retrieve the DigitalSignatureCollection count. It sets a boolean flag based on whether the count is greater than zero, indicating a digital signature, and prints the signed status while handling any runtime or reflection errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve digital signature count via reflection
            int signatureCount = 0;
            try
            {
                PropertyInfo dsProp = workbook.GetType().GetProperty("DigitalSignatureCollection", BindingFlags.Public | BindingFlags.Instance);
                if (dsProp != null)
                {
                    object dsCollection = dsProp.GetValue(workbook);
                    if (dsCollection != null)
                    {
                        PropertyInfo countProp = dsCollection.GetType().GetProperty("Count", BindingFlags.Public | BindingFlags.Instance);
                        if (countProp != null)
                        {
                            signatureCount = (int)countProp.GetValue(dsCollection);
                        }
                    }
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine($"Reflection error while checking signatures: {reflEx.Message}");
            }

            bool isSigned = signatureCount > 0;

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
