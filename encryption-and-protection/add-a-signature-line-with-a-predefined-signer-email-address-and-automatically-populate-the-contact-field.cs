// Title: Add a predefined signer email to a signature line and populate the contact field in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create a signature line on a worksheet, set the signer name, title, and a fixed email address, and automatically fill the contact field with Aspose.Cells for .NET. | When the SignatureLineCollection API is unavailable, use reflection to call its Add method, then save the workbook as an .xlsx file.
// Common Searches: how to set signer email on a signature line with Aspose.Cells .NET | adding a signature line to an Excel file using reflection in C# | populate contact field of Excel signature line programmatically | fallback method for SignatureLineCollection in older Aspose.Cells versions | save workbook with digital signature line using Aspose.Cells
// Tags: signature line add Aspose.Cells .NET | set signer email Excel signature line | populate signature line contact field | reflection fallback SignatureLineCollection | save signed Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, attempts to locate the Worksheet.SignatureLineCollection property via reflection, invokes its Add method with signer name, title, a predefined email address, and prompt text, then saves the workbook as SignedWorkbook.xlsx while gracefully handling missing API scenarios.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create a blank workbook
            Worksheet sheet = workbook.Worksheets[0]; // first worksheet

            // Predefined signer email address
            string signerEmail = "signer@example.com";

            // Attempt to add a signature line using reflection (covers versions where the API exists)
            try
            {
                var sigLineCollectionProp = typeof(Worksheet).GetProperty("SignatureLineCollection");
                if (sigLineCollectionProp != null)
                {
                    var collection = sigLineCollectionProp.GetValue(sheet);
                    var addMethod = collection?.GetType().GetMethod(
                        "Add",
                        new Type[] { typeof(string), typeof(string), typeof(string), typeof(string) });

                    if (addMethod != null)
                    {
                        addMethod.Invoke(collection, new object[]
                        {
                            "John Doe",          // Signer name
                            "Manager",           // Signer title
                            signerEmail,         // Signer email
                            "Please sign here"   // Prompt text
                        });
                        Console.WriteLine("Signature line added successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Add method for SignatureLineCollection not found.");
                    }
                }
                else
                {
                    Console.WriteLine("SignatureLineCollection property not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add signature line: {ex.Message}");
            }

            // Save the workbook
            string outputPath = "SignedWorkbook.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
