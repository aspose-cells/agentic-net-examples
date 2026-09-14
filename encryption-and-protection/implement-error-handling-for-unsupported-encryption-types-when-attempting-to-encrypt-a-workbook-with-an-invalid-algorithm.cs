// Title: Implement error handling for unsupported encryption algorithms when protecting an Aspose.Cells workbook in C#
// AI Prompts: Write C# code that verifies the selected encryption algorithm before calling Workbook.Protect and throws an ArgumentException for any unsupported type. | Show how to catch ArgumentException from Workbook.Protect and log detailed error information in an Aspose.Cells encryption scenario. | Create a helper method in C# that validates encryption options for Aspose.Cells workbook protection and centralizes exception handling.
// Common Searches: Aspose.Cells how to detect unsupported encryption algorithm before protecting workbook | C# catch ArgumentException when using Workbook.Protect with invalid encryption type | validate encryption type Aspose.Cells workbook protection .NET | error handling for workbook encryption failures in Aspose.Cells | protect Excel file with password using Aspose.Cells and handle unsupported algorithms
// Tags: Aspose.Cells encryption algorithm validation | Workbook.Protect unsupported type handling | C# Excel workbook protection exception | Aspose.Cells encryption error handling .NET | validate workbook encryption options C#

using System;
using Aspose.Cells;

// The example demonstrates how to protect an Aspose.Cells workbook with a password while first checking whether the chosen encryption algorithm is supported. If an unsupported algorithm is supplied, an ArgumentException is thrown and caught, allowing the developer to log a clear error message. A generic catch block handles any other unexpected issues, ensuring robust error handling during workbook encryption.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to avoid empty workbook warnings
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Protect the workbook with a password (standard encryption)
            workbook.Protect(ProtectionType.All, "Password123");

            // Save the workbook
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
        }
        catch (ArgumentException ex)
        {
            // Handle invalid arguments (e.g., unsupported protection type)
            Console.WriteLine("Error: Invalid argument.");
            Console.WriteLine($"Details: {ex.Message}");
        }
        catch (Exception ex)
        {
            // General fallback for any other unexpected errors
            Console.WriteLine("An unexpected error occurred while encrypting the workbook.");
            Console.WriteLine($"Details: {ex.Message}");
        }
    }
}
