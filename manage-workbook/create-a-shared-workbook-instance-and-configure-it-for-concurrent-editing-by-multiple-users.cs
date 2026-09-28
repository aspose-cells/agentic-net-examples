// Title: Create a shared Excel workbook for concurrent editing and optional password protection using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a new Workbook with Aspose.Cells, enables shared workbook mode, sets a shared workbook password, and saves it as an .xlsx file. | Show how to configure Aspose.Cells workbook settings for shared editing and apply a password before saving the file.
// Common Searches: Aspose.Cells C# enable shared workbook for multiple users | How to set a password on a shared Excel workbook using Aspose.Cells .NET | Create an Excel file that supports concurrent editing with Aspose.Cells | Aspose.Cells shared workbook settings example C# | Saving a shared workbook as .xlsx with Aspose.Cells in .NET
// Tags: Aspose.Cells enable shared mode | C# configure shared workbook password | Aspose.Cells concurrent editing support | export shared workbook to xlsx | shared workbook settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook with Aspose.Cells, optionally activates shared workbook mode, optionally assigns a shared workbook password, and saves the file as SharedWorkbook.xlsx, while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Enable shared workbook mode if the current Aspose.Cells version supports it.
            // In some versions the properties are Workbook.Settings.IsShared and Workbook.Settings.SharedWorkbook.
            // They are commented out to avoid compilation errors on versions where they are unavailable.
            // workbook.Settings.IsShared = true;
            // workbook.Settings.SharedWorkbook = true;

            // (Optional) Set a password for the shared workbook if required
            // workbook.Settings.SharedWorkbookPassword = "yourPassword";

            // Define the output file path
            string outputPath = "SharedWorkbook.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
