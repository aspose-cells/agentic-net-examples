using System;
using Aspose.Cells;

class SaveSignedWorkbook
{
    static void Main()
    {
        // Path to the existing digitally signed workbook
        string sourcePath = "SignedWorkbook.xlsx";

        // Path where the copy will be saved
        string destinationPath = "SignedWorkbookCopy.xlsx";

        // Load the signed workbook from disk
        Workbook workbook = new Workbook(sourcePath);

        // Optional: check if the workbook is digitally signed
        if (workbook.IsDigitallySigned)
        {
            Console.WriteLine("The workbook is digitally signed. Proceeding to save a copy.");
        }
        else
        {
            Console.WriteLine("Warning: The workbook does not contain a digital signature.");
        }

        // Save the workbook to a new location.
        // The digital signature is preserved automatically during the save operation.
        workbook.Save(destinationPath);

        Console.WriteLine($"Workbook successfully saved to: {destinationPath}");
    }
}