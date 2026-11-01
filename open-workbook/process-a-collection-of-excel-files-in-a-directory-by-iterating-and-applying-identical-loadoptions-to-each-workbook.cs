// Title: Load each .xlsx workbook from a folder with the same LoadOptions and overwrite using Aspose.Cells for .NET
// Common Searches: C# load every .xlsx file in a folder with a predefined LoadOptions using Aspose.Cells | batch apply password LoadOptions to multiple Excel workbooks in Aspose.Cells .NET | iterate through Excel files in a directory and save changes back with Aspose.Cells | how to use Aspose.Cells LoadOptions for bulk processing of Excel files in C#
// Tags: batch load workbooks with Aspose.Cells LoadOptions | apply identical LoadOptions to multiple .xlsx files | overwrite original Excel workbook after processing | set password and encoding in Aspose.Cells LoadOptions | directory iteration for Excel files using Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example enumerates all .xlsx files in a specified folder, loads each workbook with a shared LoadOptions object (including format, optional password, and encoding), provides a placeholder for custom processing, and then saves the workbook back to its original location, overwriting the file.
class Program
{
    static void Main()
    {
        // Path to the directory containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        // Define the LoadOptions that will be applied to every workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        // Example of additional settings (uncomment if needed)
        // loadOptions.Password = "yourPassword";
        // loadOptions.Encoding = Encoding.UTF8;

        // Iterate over all .xlsx files in the specified directory
        foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
        {
            // Load the workbook using the predefined LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // -------------------------------------------------
            // Insert any workbook processing logic here
            // -------------------------------------------------

            // Save the workbook back to the same file (overwrites original)
            workbook.Save(filePath);
        }
    }
}
