using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionCheck
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "sample.xlsx";
            const string outputPath = "sample_checked.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the protection settings for the worksheet
                Protection protection = worksheet.Protection;

                // Check whether editing of drawing objects is allowed on the protected worksheet
                bool allowEditingObjects = protection.AllowEditingObject;

                // Log the compliance result
                Console.WriteLine($"Worksheet \"{worksheet.Name}\" allows editing objects: {allowEditingObjects}");

                // Save the workbook (optional, no changes made to protection settings)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}