// Title: Add a signature line shape to an Excel workbook and fetch a certificate by thumbprint from the Windows certificate store with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to place a rectangular shape labeled "Signature Line" in a worksheet and then obtains an X509Certificate2 from the current user's personal store using a given thumbprint. | Create a reusable C# method that searches the Windows certificate store for a specific thumbprint and demonstrate how to associate the found certificate with a placeholder signature shape in an Excel file saved via Aspose.Cells.
// Common Searches: how to add a rectangle shape as a signature line in an Excel file using Aspose.Cells C# | retrieve a Windows certificate by thumbprint in C# for Excel digital signing | Aspose.Cells example linking a signature placeholder to an X509 certificate | C# code to open the personal certificate store and find a certificate with a specific thumbprint | save an Excel workbook with a signature line shape using Aspose.Cells .NET
// Tags: Aspose.Cells insert signature shape | C# retrieve certificate by thumbprint | Windows personal store X509 lookup | Excel digital signature placeholder Aspose | save workbook with shape Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook with Aspose.Cells, adds a rectangular shape labeled "Signature Line" as a placeholder, retrieves an X509Certificate2 from the current user's personal certificate store using a supplied thumbprint, and saves the workbook to a file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            try
            {
                // Add a simple rectangle shape to act as a placeholder for a signature line
                // Parameters: shape type, upper left row, upper left column, upper left row offset (pixels),
                // upper left column offset (pixels), width (pixels), height (pixels)
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // correct enum for shape type
                    2, 2, 0, 0, 200, 50);
                shape.Text = "Signature Line";
            }
            catch (Exception shapeEx)
            {
                Console.WriteLine($"Shape error: {shapeEx.Message}");
            }

            // OPTIONAL: Retrieve a certificate (not used further in this simplified example)
            try
            {
                string thumbprint = "ABCD1234EF567890ABCD1234EF567890ABCD1234"; // replace with actual thumbprint
                X509Certificate2 certificate = GetCertificateByThumbprint(thumbprint);
                // Certificate can be used for further processing if needed
            }
            catch (Exception certEx)
            {
                Console.WriteLine($"Certificate error: {certEx.Message}");
            }

            // Save the workbook
            string outputPath = "SignedWorkbook.xlsx";

            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Save error: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static X509Certificate2 GetCertificateByThumbprint(string thumbprint)
    {
        // Clean the thumbprint string (remove spaces, line breaks, etc.)
        string cleanThumbprint = thumbprint.Replace(" ", "")
                                           .Replace("\t", "")
                                           .Replace("\r", "")
                                           .Replace("\n", "");

        // Open the personal certificate store of the current user
        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);

            // Find the certificate by thumbprint
            X509Certificate2Collection certs = store.Certificates.Find(
                X509FindType.FindByThumbprint, cleanThumbprint, false);

            if (certs.Count == 0)
                throw new Exception($"Certificate with thumbprint {thumbprint} not found.");

            return certs[0];
        }
    }
}
