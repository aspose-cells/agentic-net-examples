// Title: Detect plain‑text passwords in Excel workbook custom document properties using Aspose.Cells for .NET
// Common Searches: C# code to scan Excel custom document properties for stored passwords with Aspose.Cells | how to identify plain text password values in workbook metadata using Aspose.Cells .NET | raise an alert when a password is found in Excel custom properties in a .NET application | validate that Excel files do not contain insecure password entries in custom metadata using Aspose.Cells | detect insecure password storage in Excel file custom properties programmatically
// Tags: detect plain text password in Excel custom document properties Aspose.Cells | scan workbook metadata for insecure password storage .NET | validate Excel custom properties security using Aspose.Cells | C# check custom document properties for password leakage | Aspose.Cells workbook metadata password audit

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through its CustomDocumentProperties, and raises an alert for any property whose name contains "password" and whose value is a non‑empty string, helping developers identify plain‑text passwords stored in workbook metadata while handling missing files and other exceptions.
class Program
{
    static void Main()
    {
        // Path to the workbook to be inspected
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file '{filePath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied; assumes the file is not encrypted for reading)
            Workbook workbook = new Workbook(filePath);

            bool plainTextPasswordFound = false;

            // Examine each custom document property
            foreach (var prop in workbook.CustomDocumentProperties)
            {
                // Identify properties whose name suggests they may hold a password
                if (prop.Name.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Verify the value is a non‑empty string
                    if (prop.Value is string value && !string.IsNullOrWhiteSpace(value))
                    {
                        plainTextPasswordFound = true;
                        Console.WriteLine($"Alert: Plain text password detected in custom metadata property '{prop.Name}'.");
                    }
                }
            }

            if (!plainTextPasswordFound)
            {
                Console.WriteLine("No plain text password found in custom metadata.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., file format issues, permission problems)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
