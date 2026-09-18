// Title: Remove all signature lines from an Excel workbook and insert a new signature line with custom signer information using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, iterates through each worksheet, deletes every signature line via the SignatureLineCollection, and then adds a new signature line with specified signer name, title, email, and comments. | Show how to use reflection to access the SignatureLineCollection property in Aspose.Cells when the API is not directly exposed, ensuring compatibility across different library versions while removing and adding signature lines.
// Common Searches: C# Aspose.Cells delete all signature lines from an existing Excel file | How to add a new digital signature line with custom signer details using Aspose.Cells .NET | Using reflection to manipulate SignatureLineCollection in older Aspose.Cells versions | Replace existing Excel signature lines with updated signer information in .NET
// Tags: Aspose.Cells remove signature lines | Aspose.Cells add signature line | SignatureLineCollection reflection .NET | Excel digital signature manipulation Aspose | update signer details in workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an input.xlsx workbook with Aspose.Cells, checks each worksheet for a SignatureLineCollection (using reflection for version compatibility), removes all existing signature lines, adds a new signature line on the first sheet with custom signer name, title, email, and comments, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Attempt to work with signature lines via reflection/dynamic (for compatibility with different Aspose.Cells versions)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var sigProp = sheet.GetType().GetProperty("SignatureLineCollection");
                if (sigProp == null)
                {
                    // SignatureLineCollection not supported in this version
                    continue;
                }

                dynamic sigCollection = sigProp.GetValue(sheet);
                if (sigCollection == null)
                    continue;

                // Remove existing signature lines safely
                try
                {
                    int count = sigCollection.Count;
                    for (int i = count - 1; i >= 0; i--)
                    {
                        sigCollection.RemoveAt(i);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to clear existing signatures: {ex.Message}");
                }

                // Add a new signature line (only on the first worksheet for demonstration)
                if (sheet.Index == 0)
                {
                    try
                    {
                        // Define the cell range that will contain the signature line (zero‑based indices)
                        var newSignature = sigCollection.Add(5, 2, 7, 5);

                        // Set signer information
                        newSignature.Signer = "John Doe";
                        newSignature.SignerTitle = "Chief Financial Officer";
                        newSignature.Email = "john.doe@example.com";
                        newSignature.SignerComments = "Approved financial report";

                        // Optional display settings
                        newSignature.AllowComments = true;
                        newSignature.ShowSignDate = true;
                        newSignature.ShowSignTime = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to add a new signature line: {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
