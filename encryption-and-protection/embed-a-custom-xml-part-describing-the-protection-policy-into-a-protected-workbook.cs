// Title: How to embed a UTF-8 custom XML part that defines a protection policy into a password-protected Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new Workbook, applies password protection, converts a protection‑policy XML string to UTF-8 bytes, adds it as a custom XML part named 'ProtectionPolicy' with CustomXmlParts.Add, and saves the file as an .xlsx. | Show how to embed metadata describing allowed actions (Read, Print) into a protected Excel file by using the Aspose.Cells CustomXmlParts API with byte‑array parameters.
// Common Searches: aspnet add custom xml part to password protected Excel workbook using Aspose.Cells | how to store protection policy XML inside an encrypted .xlsx with Aspose.Cells | Aspose.Cells CustomXmlParts.Add example for embedding metadata in protected workbook | C# embed protection policy as custom XML in protected Excel file | save workbook with password and custom XML part Aspose.Cells .NET
// Tags: customxmlparts add utf-8 byte array | password protect workbook Aspose.Cells | embed protectionpolicy xml in Excel | save workbook with embedded custom xml | Aspose.Cells metadata embedding example

using Aspose.Cells;
using System;
using System.Text;

// // Creates a workbook, adds sample data, protects it with a password, converts a protection‑policy XML string to UTF‑8 bytes, embeds it as a custom XML part named 'ProtectionPolicy' using CustomXmlParts.Add, and saves the result as a protected .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B1"].PutValue(100);

            // Protect the workbook with a password
            workbook.Protect(ProtectionType.All, "StrongPassword123");

            // Define the custom XML that describes the protection policy
            string protectionPolicyXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<ProtectionPolicy>
    <Password>StrongPassword123</Password>
    <AllowedActions>
        <Action>Read</Action>
        <Action>Print</Action>
    </AllowedActions>
</ProtectionPolicy>";

            // Convert XML string to UTF‑8 byte array (required by CustomXmlParts.Add)
            byte[] xmlBytes = Encoding.UTF8.GetBytes(protectionPolicyXml);
            // Convert the part name to a byte array as required by the API overload
            byte[] partNameBytes = Encoding.UTF8.GetBytes("ProtectionPolicy");

            // Embed the custom XML part into the workbook (argument order: name, data)
            workbook.CustomXmlParts.Add(partNameBytes, xmlBytes);

            // Save the protected workbook with the embedded XML part
            workbook.Save("ProtectedWorkbookWithPolicy.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
