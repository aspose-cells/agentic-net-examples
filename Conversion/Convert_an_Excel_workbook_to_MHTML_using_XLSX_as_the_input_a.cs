using System;
using Aspose.Cells; // Aspose.Cells namespace

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired path for the MHTML output
        string destPath = "output.mht";

        // Load the workbook from the XLSX file (create & load)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as MHTML (save)
        workbook.Save(destPath, SaveFormat.MHtml);

        Console.WriteLine("Conversion from XLSX to MHTML completed successfully.");
    }
}