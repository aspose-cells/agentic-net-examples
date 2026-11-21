// Title: Preserve the original thread culture while loading an Excel workbook with French CultureInfo using Aspose.Cells LoadOptions in C#
// Common Searches: c# aspocells load workbook with french cultureinfo and keep original thread culture | how to set CultureInfo on LoadOptions for Aspose.Cells and restore thread culture afterwards | aspocells create fallback workbook when file is missing in c# example | restore original thread culture after loading an Excel file with Aspose.Cells LoadOptions | load an xlsx file with French locale using Aspose.Cells without affecting current thread culture
// Tags: load workbook with CultureInfo Aspose.Cells | preserve thread culture Aspose.Cells | fallback workbook creation Aspose.Cells | French locale loading Excel Aspose.Cells | restore original thread culture after workbook load

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example captures the current thread's CultureInfo, applies a French CultureInfo to LoadOptions, loads an Excel workbook with Aspose.Cells, creates a simple fallback workbook if the file is missing, and finally restores the original thread culture before accessing cell values.
    class Program
    {
        static void Main()
        {
            // Preserve the original thread culture
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

            // Define French culture for loading the workbook
            CultureInfo frenchCulture = new CultureInfo("fr-FR");

            // Prepare load options with French culture
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = frenchCulture
            };

            Workbook workbook = null;
            string filePath = "SampleFrench.xlsx";

            try
            {
                // Ensure the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    // If the file is missing, create a simple workbook as a fallback
                    workbook = new Workbook();
                    workbook.Worksheets[0].Cells["A1"].PutValue("Fallback");
                    workbook.Save(filePath);
                    Console.WriteLine($"File not found. Created a fallback workbook at '{Path.GetFullPath(filePath)}'.");
                }
                else
                {
                    // Load the workbook using the specified LoadOptions
                    workbook = new Workbook(filePath, loadOptions);
                }

                // Restore the original thread culture after loading
                Thread.CurrentThread.CurrentCulture = originalCulture;

                // Verify that the culture has been restored
                Console.WriteLine("Current thread culture after restore: " + Thread.CurrentThread.CurrentCulture.Name);

                // Example usage of the loaded workbook
                Console.WriteLine("Value of A1: " + workbook.Worksheets[0].Cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                // Ensure the original culture is restored even if an exception occurs
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }
        }
    }
}
