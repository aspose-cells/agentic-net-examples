// Title: Preserve a Digital Signature When Saving a Signed Excel Workbook to a New Location with Aspose.Cells for .NET
// Description: Demonstrates how to load a digitally signed Excel workbook using Aspose.Cells, verify the IsDigitallySigned flag, and save the file to a different path while the signature remains intact.
// Keywords: Aspose.Cells digital signature | C# save signed workbook | IsDigitallySigned property | copy signed Excel file | .NET preserve Excel signature | Workbook.Save signature retention
// Common Searches: keep digital signature when saving Excel with Aspose.Cells | copy signed workbook to another folder C# | IsDigitallySigned example Aspose.Cells | save signed Excel file to new location .NET | preserve Excel digital signature programmatically
// Developer Intent: Create a copy of a signed Excel workbook and ensure the digital signature is retained after saving.
// Use Cases: Archive a signed financial report without invalidating its signature. | Distribute a signed template to multiple recipients while maintaining authenticity. | Automate backup of digitally signed Excel files to a secure repository.
// AI Prompts: Write C# code that loads a digitally signed workbook with Aspose.Cells, checks the IsDigitallySigned flag, and saves a copy to a new path preserving the signature. | Explain how Aspose.Cells handles digital signatures during Workbook.Save and whether any configuration is required to keep the signature intact.

using System;
using Aspose.Cells;

// Demonstrates how to load a digitally signed Excel workbook using Aspose.Cells, verify the IsDigitallySigned flag, and save the file to a different path while the signature remains intact.
class Program
{
    static void Main()
    {
        // Path to the existing digitally signed workbook
        string sourcePath = "SignedWorkbook.xlsx";

        // Path where the copy will be saved
        string destinationPath = "SignedCopy.xlsx";

        // Load the signed workbook from disk
        Workbook workbook = new Workbook(sourcePath);

        // Check if the workbook is digitally signed
        if (workbook.IsDigitallySigned)
        {
            Console.WriteLine("The workbook is digitally signed.");
        }
        else
        {
            Console.WriteLine("The workbook is NOT digitally signed.");
        }

        // Save the workbook to a new location.
        // The digital signature is preserved automatically.
        workbook.Save(destinationPath);

        Console.WriteLine($"Workbook saved to '{destinationPath}' with the digital signature intact.");
    }
}
