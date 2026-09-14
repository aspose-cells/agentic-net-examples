// Title: Generate HTML from a C# Aspose.Cells workbook with list‑validation dropdowns and verify the <select> element appears
// AI Prompts: Write C# code that uses Aspose.Cells to add a list‑type data validation to a cell range, save the workbook as HTML, and programmatically confirm the output contains a <select> tag. | Show how to create a dropdown list in cells C1:C5 referencing A1:A3, export the sheet to HTML, and assert that the generated HTML includes a <select> element.
// Common Searches: Aspose.Cells C# export worksheet with data validation list to HTML | how to render Excel dropdown as <select> in HTML using Aspose.Cells | verify dropdown rendering in Aspose.Cells HTML output C# | save workbook with list validation as HTML file Aspose.Cells .NET | C# example for data validation dropdown to HTML conversion Aspose.Cells
// Tags: Aspose.Cells list validation HTML export | C# data validation dropdown to HTML | render Excel dropdown as select tag | verify generated HTML contains select element | SaveFormat.Html workbook conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDataValidationHtmlDemo
{
    // Demonstrates creating a workbook, adding a list‑type data validation to cells C1:C5 referencing A1:A3, saving the workbook as HTML with Aspose.Cells, and checking the resulting file for a <select> element to confirm proper dropdown rendering.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a list of values that will be used for the dropdown (e.g., A1:A3)
                sheet.Cells["A1"].PutValue("Option 1");
                sheet.Cells["A2"].PutValue("Option 2");
                sheet.Cells["A3"].PutValue("Option 3");

                // Define the range where the dropdown should appear (e.g., C1:C5)
                // Use CellArea to specify the target cells
                CellArea dropdownArea = new CellArea
                {
                    StartRow = 0,   // C1 (row 0)
                    EndRow = 4,     // C5 (row 4)
                    StartColumn = 2, // column C (index 2)
                    EndColumn = 2
                };

                // Add a list validation for the specified area
                int validationIndex = sheet.Validations.Add(dropdownArea);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.List;               // set validation type
                validation.Formula1 = "=$A$1:$A$3";                  // source list range
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Selection";
                validation.ErrorMessage = "Please select a value from the list.";

                // Prepare output HTML path
                string htmlPath = "WorkbookWithDropdown.html";
                string htmlDir = Path.GetDirectoryName(htmlPath);
                if (!string.IsNullOrEmpty(htmlDir) && !Directory.Exists(htmlDir))
                {
                    Directory.CreateDirectory(htmlDir);
                }

                // Save the workbook as HTML
                workbook.Save(htmlPath, SaveFormat.Html);

                // Verify that the generated HTML contains a <select> element (Aspose.Cells renders dropdowns as <select>)
                if (File.Exists(htmlPath))
                {
                    string htmlContent = File.ReadAllText(htmlPath);
                    if (htmlContent.IndexOf("<select", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine("Verification succeeded: <select> element found in the HTML output.");
                    }
                    else
                    {
                        Console.WriteLine("Verification failed: <select> element not found in the HTML output.");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
