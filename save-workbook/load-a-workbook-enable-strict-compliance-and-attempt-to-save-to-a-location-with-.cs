using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with a valid file path)
        string inputPath = "input.xlsx";
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Enable ISO/IEC 29500:2008 Strict compliance for OOXML
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Attempt to save the workbook to a location that typically requires elevated permissions
        string protectedPath = @"C:\Windows\System32\protected.xlsx";

        try
        {
            workbook.Save(protectedPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save workbook: " + ex.Message);
        }
    }
}