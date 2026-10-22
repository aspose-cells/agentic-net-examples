// Title: Identify plain‑text encryption passwords hidden in Excel custom document properties with Aspose.Cells for .NET
// Common Searches: C# code to scan Excel custom document properties for stored passwords using Aspose.Cells | how to detect insecure plain text password in an .xlsx file with Aspose.Cells .NET | raise an alert when a workbook's encryption password is saved in custom metadata in C# | check Excel workbook metadata for clear‑text password using Aspose.Cells library | validate that Excel custom properties do not contain clear‑text passwords in .NET
// Tags: custom document properties scanning Aspose.Cells | password exposure detection in Excel workbooks | metadata security validation C# Aspose.Cells | encryption password plain‑text alert .NET | secure workbook metadata check

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through all custom document properties, and flags any property whose name contains "password" and whose value is a non‑empty string, emitting an alert for plain‑text passwords or confirming none were found.
class WorkbookPasswordChecker
{
    static void Main(string[] args)
    {
        // Path to the workbook to be inspected
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file '{filePath}' was not found.");
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

        // Flag to indicate if a plain‑text password was found
        bool passwordFound = false;

        // Iterate through all custom document properties
        foreach (var prop in workbook.CustomDocumentProperties)
        {
            // Check if the property name suggests it may hold a password
            if (prop.Name.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Ensure the value is a non‑empty string
                if (prop.Value is string strValue && !string.IsNullOrWhiteSpace(strValue))
                {
                    // Alert: plain‑text password detected
                    Console.WriteLine($"Alert: Encryption password stored in plain text in custom metadata property '{prop.Name}'.");
                    passwordFound = true;
                }
            }
        }

        if (!passwordFound)
        {
            Console.WriteLine("No plain‑text encryption password found in custom metadata.");
        }
    }
}
