// Title: Add a password hint as a custom document property after protecting workbook structure with Aspose.Cells for .NET (C#)
// Common Searches: how to store a password hint in an Excel file using Aspose.Cells C# | protect only the structure of a workbook and add a custom property with Aspose.Cells | Aspose.Cells add custom document property after workbook protection example | C# code to protect workbook structure and include a password hint comment | using Aspose.Cells to add a PasswordHint custom property to a protected Excel workbook
// Tags: protect workbook structure Aspose.Cells C# | add custom document property Aspose.Cells | password hint custom property Excel | Aspose.Cells workbook protection with hint | Excel custom property for password hint .NET

using Aspose.Cells;
using System;

// The example creates a new Workbook, protects its structure with a password via workbook.Protect(ProtectionType.Structure, password), adds a custom document property named "PasswordHint" containing a hint string, and saves the protected workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook's structure with a password
            string password = "MySecretPassword";
            // Use the correct API to protect only the structure
            workbook.Protect(ProtectionType.Structure, password);

            // Add a custom property to serve as a password hint comment
            // The custom property is stored in the workbook's CustomDocumentProperties collection
            string hint = "The password is the name of my first pet.";
            workbook.CustomDocumentProperties.Add("PasswordHint", hint);

            // Save the workbook to a file
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
