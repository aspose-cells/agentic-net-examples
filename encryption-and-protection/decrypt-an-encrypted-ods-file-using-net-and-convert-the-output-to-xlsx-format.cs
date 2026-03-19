using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class DecryptOdsAndConvertToXlsx
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string sourcePath = "encrypted_input.ods";
            string password = "myPassword";
            string destPath = "output.xlsx";

            try
            {
                OdsLoadOptions loadOptions = new OdsLoadOptions
                {
                    Password = password
                };

                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

                ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                Console.WriteLine($"Decryption and conversion completed successfully. Output file: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during decryption/conversion: {ex.Message}");
            }
        }
    }
}