// Title: Attempt to detect Excel workbook encryption algorithm and evaluate its strength using Aspose.Cells for .NET (C#)
// Common Searches: how to load a password‑protected XLSX file with Aspose.Cells in C# | does Aspose.Cells expose the encryption algorithm of an encrypted Excel workbook | determine if an Excel file uses strong or weak encryption with Aspose.Cells .NET | retrieve encryption type (AES‑128, AES‑256, etc.) of an XLSX file using C# | check workbook encryption strength after opening with Aspose.Cells LoadOptions
// Tags: Aspose.Cells workbook encryption detection | C# load password protected XLSX | evaluate Excel encryption strength .NET | Aspose.Cells LoadOptions password usage | identify Excel file encryption algorithm | assess weak vs strong Excel encryption

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to open a password‑protected Excel workbook using Aspose.Cells LoadOptions in C#. While the workbook can be opened successfully, Aspose.Cells does not provide a property to reveal the underlying encryption algorithm, so the code can only confirm access, not classify the encryption as strong or weak.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Ensure the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        // Password required to open the workbook (if known)
        string password = "yourPassword";

        try
        {
            // Load the workbook with password using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Open the workbook
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Since Aspose.Cells does not expose the encryption algorithm after loading,
            // we simply confirm that the workbook was opened successfully.
            Console.WriteLine("Workbook opened successfully with the provided password.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., incorrect password, corrupted file)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
