// Title: Create an HTML file with an in‑cell dropdown list using Aspose.Cells for .NET and confirm the <select> tag is generated
// Common Searches: Aspose.Cells C# export worksheet with data validation list to HTML | how to generate a <select> element from Excel data validation using Aspose.Cells | verify dropdown rendering in HTML output from an Aspose.Cells workbook | C# code to add a list validation and save as HTML with Aspose.Cells | check if Aspose.Cells HTML export includes a select tag for in‑cell dropdowns
// Tags: Aspose.Cells export list validation to HTML | C# generate select element from Excel validation | HtmlSaveOptions dropdown rendering | validation.InCellDropDown HTML select conversion | Aspose.Cells verify HTML dropdown presence

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataValidationHtmlDemo
{
    // The example creates a workbook, adds a worksheet named "Demo", places a label in cell A1 and an empty cell B1, applies a list‑type data validation with three options to B1, enables the in‑cell dropdown, saves the workbook as HTML using HtmlSaveOptions, reads the generated file, and checks for a <select> element to confirm that the dropdown is rendered in the HTML output.
    class Program
    {
        static void Main(string[] args)
        {
            // Define paths for the generated HTML file.
            string htmlPath = Path.Combine(Path.GetTempPath(), "DataValidationDemo.html");

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and add a worksheet.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Demo";

                // ------------------------------------------------------------
                // 2. Populate some sample data (optional, just for context).
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Select an option:");
                sheet.Cells["B1"].PutValue(string.Empty); // cell that will hold the dropdown

                // ------------------------------------------------------------
                // 3. Add a data‑validation list to cell B1.
                // ------------------------------------------------------------
                // Define the range that will receive the validation.
                CellArea validationArea = CellArea.CreateCellArea("B1", "B1");

                // Access the worksheet's validation collection and add a new validation.
                ValidationCollection validations = sheet.Validations;
                int validationIndex = validations.Add(validationArea);
                Validation validation = validations[validationIndex];

                // Configure the validation as a list with three options.
                validation.Type = ValidationType.List;
                // The list items are supplied as a comma‑separated string.
                validation.Formula1 = "Option1,Option2,Option3";

                // Enable the in‑cell drop‑down (this is what causes a <select> in HTML).
                validation.InCellDropDown = true;

                // Optional: make the validation visible to the user.
                validation.ShowInput = true;
                validation.InputMessage = "Choose an option from the list.";
                validation.ErrorMessage = "Invalid selection.";

                // ------------------------------------------------------------
                // 4. Save the workbook as HTML.
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");

                // ------------------------------------------------------------
                // 5. Verify that the generated HTML contains a <select> element.
                // ------------------------------------------------------------
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    bool containsSelect = htmlContent.IndexOf("<select", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (containsSelect)
                    {
                        Console.WriteLine("Verification succeeded: <select> element found in the HTML output.");
                    }
                    else
                    {
                        Console.WriteLine("Verification failed: <select> element NOT found in the HTML output.");
                    }
                }
                else
                {
                    Console.WriteLine("Verification failed: HTML file was not created.");
                }

                // ------------------------------------------------------------
                // 6. (Optional) Clean up temporary files.
                // ------------------------------------------------------------
                // Uncomment the following line if you want the demo to delete the HTML file after execution.
                // File.Delete(htmlPath);
            }
            catch (Exception ex)
            {
                // Report any unexpected errors.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
