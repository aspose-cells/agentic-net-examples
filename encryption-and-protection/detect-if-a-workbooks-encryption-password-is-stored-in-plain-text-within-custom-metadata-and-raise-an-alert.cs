// Title: Detect plain‑text encryption passwords hidden in custom document properties of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates over a workbook's CustomDocumentProperties and logs an alert when a property name contains the word "password" and its value is a non‑empty string. | Create a reusable method in Aspose.Cells that returns true if any custom metadata field in an Excel file stores an unencrypted password, and writes the offending property name to the console. | Modify the sample to collect all matching custom properties, output their names and values, and exit the application with a non‑zero status code if any unencrypted passwords are found.
// Common Searches: how to audit Excel custom document properties for stored passwords using Aspose.Cells in C# | Aspose.Cells detect password stored in workbook metadata .NET | scan .xlsx custom metadata for encryption credentials programmatically | raise warning when Excel file contains password in custom properties Aspose.Cells | validate that Excel workbook does not expose encryption password in custom fields
// Tags: scan custom document properties Aspose.Cells | password exposure Excel | workbook metadata security validation .NET | Aspose.Cells encryption password audit | Excel file custom metadata check

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, iterates through its custom document properties, and prints an alert if any property name includes "password" and its value is a non‑empty string; otherwise it reports that no plain‑text encryption password was found.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found at path '{filePath}'.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Scan custom document properties for plain‑text passwords
        bool passwordDetected = false;
        foreach (var prop in workbook.CustomDocumentProperties)
        {
            // Identify properties whose name suggests a password
            if (prop.Name != null && prop.Name.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Verify the value is a non‑empty string
                if (prop.Value is string value && !string.IsNullOrWhiteSpace(value))
                {
                    Console.WriteLine($"Alert: Encryption password stored in plain text in custom metadata property '{prop.Name}'.");
                    passwordDetected = true;
                }
            }
        }

        if (!passwordDetected)
        {
            Console.WriteLine("No plain‑text encryption password found in custom metadata.");
        }
    }
}
