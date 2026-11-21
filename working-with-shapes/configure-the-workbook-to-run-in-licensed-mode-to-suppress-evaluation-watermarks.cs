// Title: Configure Aspose.Cells for C# to use a license file and eliminate evaluation watermarks when creating and saving a workbook
// Common Searches: how to load Aspose.Cells.lic in a C# application to remove the evaluation banner | C# code example for setting Aspose.Cells license before saving an Excel file | suppress Aspose.Cells evaluation watermark by applying a license at runtime | using Aspose.Cells License class to run in licensed mode in .NET | apply Aspose.Cells license programmatically to avoid watermarks in generated workbooks
// Tags: Aspose.Cells license activation .NET | programmatic license setting Aspose.Cells | prevent evaluation watermark in Excel output | licensed workbook generation Aspose.Cells | load Aspose.Cells.lic at runtime C#

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for an Aspose.Cells.lic file, applies it with License.SetLicense to run in licensed mode and remove evaluation watermarks, then creates a new Workbook via the provided CreateWorkbook method and saves it as LicensedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Attempt to apply Aspose.Cells license if the file exists
            string licensePath = "Aspose.Cells.lic";
            if (File.Exists(licensePath))
            {
                try
                {
                    License license = new License();
                    license.SetLicense(licensePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to set license. {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("License file not found. Running in evaluation mode.");
            }

            // Create a new workbook using the provided create rule
            Workbook workbook = CreateWorkbook(); // <-- rule placeholder for workbook creation

            // (Optional) Perform any workbook operations here

            // Save the workbook using the provided save rule
            SaveWorkbook(workbook, "LicensedWorkbook.xlsx"); // <-- rule placeholder for saving
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // The following methods represent the mandated create and save rules.
    // Their implementations are supplied by the rule set and should not be altered.

    static Workbook CreateWorkbook()
    {
        // Rule implementation will instantiate and return a new Workbook.
        return new Workbook();
    }

    static void SaveWorkbook(Workbook wb, string filePath)
    {
        // Rule implementation will handle saving the workbook to the specified path.
        wb.Save(filePath);
    }
}
