// Title: How to apply an Aspose.Cells license in C# to eliminate evaluation watermarks before creating a Workbook
// AI Prompts: Write a C# method that receives a file path, verifies the Aspose.Cells .lic file exists, and applies the license using the License class. | Create a try‑catch block in C# that loads an Aspose.Cells license from a specified location and logs success or error details to the console. | Generate sample code that sets the Aspose.Cells license before any Workbook instantiation to ensure the workbook runs in licensed mode.
// Common Searches: c# set Aspose.Cells license before workbook creation to remove evaluation watermark | how to programmatically verify Aspose.Cells license file existence in .NET | apply Aspose.Cells .lic file in console application and suppress evaluation message | Aspose.Cells licensing example for C# console projects
// Tags: Aspose.Cells License.SetLicense C# | suppress Aspose.Cells evaluation watermark | check license file existence .NET | apply license before Workbook instantiation | exception handling Aspose.Cells license loading

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLicenseDemo
{
    // Demonstrates loading an Aspose.Cells license from a given path, verifying the file exists, applying it via the License class, and handling errors so that workbooks run in licensed mode without evaluation watermarks.
    public class WorkbookLicenseConfigurator
    {
        public static void ApplyLicense()
        {
            try
            {
                // Path to the Aspose.Cells license file
                string licensePath = @"C:\Path\To\Aspose.Cells.lic";

                // Verify that the license file is present to avoid FileNotFoundException
                if (!File.Exists(licensePath))
                {
                    Console.WriteLine($"License file not found at: {licensePath}");
                    return;
                }

                // Initialize the Aspose.Cells license object and apply the license
                var license = new License();
                license.SetLicense(licensePath);
                Console.WriteLine("Aspose.Cells license applied successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors during license application
                Console.WriteLine($"Error applying Aspose.Cells license: {ex.Message}");
            }
        }
    }

    public class Program
    {
        /// <summary>
        /// Entry point of the application.
        /// </summary>
        public static void Main(string[] args)
        {
            // Apply the Aspose.Cells license before any workbook operations
            WorkbookLicenseConfigurator.ApplyLicense();

            // Additional workbook processing can be placed here
            Console.WriteLine("Program execution completed.");
        }
    }
}
