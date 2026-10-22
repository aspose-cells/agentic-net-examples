// Title: Insert a company logo at the top of an Excel worksheet using an image smart marker with Aspose.Cells for .NET (C#)
// AI Prompts: Load a PNG file into a byte array, bind it to an image smart marker named "Logo" via WorkbookDesigner.SetDataSource, then call Process to embed the logo in the template. | Swap the placeholder image in the template with the logo byte array and save the modified workbook. | Add a company logo to the header row of an Excel file programmatically using Aspose.Cells smart markers without writing the image to disk.
// Common Searches: C# Aspose.Cells how to bind a byte array to an image smart marker | replace image marker with logo in Excel template using WorkbookDesigner | insert company logo at top of worksheet using Aspose.Cells smart markers | Aspose.Cells set image data source from PNG bytes example | populate Excel template with logo using smart markers .NET
// Tags: image smart marker with WorkbookDesigner | set image data source from byte array Aspose.Cells | swap placeholder image marker in Excel template | embed company logo in Excel worksheet C# | Aspose.Cells logo insertion example

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads a template workbook, reads a PNG logo into a byte array, assigns that array to an image smart marker named "Logo" via WorkbookDesigner, processes the markers to replace the placeholder, and saves the result as a new Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string logoPath = "logo.png";
                const string resultPath = "Result.xlsx";

                // Ensure the template workbook exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Load logo bytes if the file exists
                byte[] logoBytes = null;
                if (File.Exists(logoPath))
                {
                    logoBytes = File.ReadAllBytes(logoPath);
                }
                else
                {
                    Console.WriteLine($"Logo file not found: {logoPath}. Image marker will be ignored.");
                }

                // Work with markers in the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Assign logo to the image marker named "Logo" if available
                if (logoBytes != null)
                {
                    designer.SetDataSource("Logo", logoBytes);
                }

                // Process markers and insert the image
                designer.Process();

                // Save the modified workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
