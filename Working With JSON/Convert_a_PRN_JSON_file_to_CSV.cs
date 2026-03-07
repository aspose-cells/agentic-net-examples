using System;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            PrnJsonToCsvConverter.Run();
        }
    }

    public class PrnJsonToCsvConverter
    {
        public static void Run()
        {
            string sourceFile = "input.prn";
            string outputFile = "output.csv";

            try
            {
                ConversionUtility.Convert(sourceFile, outputFile);
                Console.WriteLine($"Conversion successful: '{sourceFile}' -> '{outputFile}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}