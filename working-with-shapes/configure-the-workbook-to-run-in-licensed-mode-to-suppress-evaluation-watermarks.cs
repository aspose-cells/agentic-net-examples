// Title: Load an Aspose.Cells license in C# to run workbooks without evaluation watermarks
// Common Searches: C# code example for applying Aspose.Cells license before creating a workbook | how to suppress Aspose.Cells evaluation watermark using a .lic file in .NET | setting Aspose.Cells license programmatically in a C# console application | remove Aspose.Cells trial watermark by loading license at runtime | initialize Aspose.Cells License object in C# to enable licensed mode
// Tags: Aspose.Cells license activation C# | disable Aspose.Cells trial watermark .NET | programmatic .lic file usage Aspose.Cells | licensed mode setup Aspose.Cells API | runtime license loading Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for an Aspose.Cells.lic file, loads it with a License object to enable licensed mode (otherwise it runs in evaluation mode), creates a new Workbook, and saves it as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Load license if it exists; otherwise run in evaluation mode
            string licensePath = "Aspose.Cells.lic";
            if (File.Exists(licensePath))
            {
                License license = new License();
                license.SetLicense(licensePath);
            }
            else
            {
                Console.WriteLine($"License file not found at '{licensePath}'. Running in evaluation mode.");
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Perform any workbook operations here
            // ...

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
