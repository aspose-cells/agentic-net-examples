// Title: How to add whole-number, list, and custom data validations to XML-mapped cells in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that loads an XML file, maps its rows to a worksheet with Aspose.Cells, and applies a whole-number validation (1‑1000) to column A, a list validation (Option1‑Option2‑Option3) to column B, and a custom validation enforcing exactly five numeric characters in column C. | Show how to modify the custom validation formula for column C so that it accepts exactly six alphanumeric characters instead of five digits, using Aspose.Cells ValidationType.Custom. | Demonstrate saving the validated workbook to both XLSX and PDF formats while preserving all data‑validation rules with Aspose.Cells.
// Common Searches: aspocells c# add whole number validation to column after XML mapping | list validation dropdown options in Aspose.Cells worksheet imported from XML | custom validation formula fixed length numeric code Aspose.Cells C# example | save Aspose.Cells workbook with data validation to PDF | apply multiple data validations to columns after XML mapping using Aspose.Cells .NET
// Tags: Aspose.Cells XML mapping with data validation | C# whole-number validation Aspose.Cells | Aspose.Cells list validation dropdown | custom formula validation fixed-length numeric code | save validated workbook as PDF Aspose.Cells

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsDataValidationExample
{
    // The example creates a new workbook, loads rows from Data.xml, writes values to columns A‑C starting at row 2, and adds three validations: whole-number between 1 and 1000 for column A, a predefined list for column B, and a custom formula ensuring exactly five numeric characters for column C. The workbook is then saved as ValidatedData.xlsx with error handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Load XML data
                string xmlPath = "Data.xml";
                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine($"XML file not found: {xmlPath}");
                    return;
                }

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlPath);

                // Map XML values to worksheet cells
                XmlNodeList rowNodes = xmlDoc.SelectNodes("//Rows/Row");
                int startRowIndex = 1; // Row 2 in Excel (zero‑based)
                int currentRow = startRowIndex;

                if (rowNodes != null)
                {
                    foreach (XmlNode rowNode in rowNodes)
                    {
                        // Column A - numeric value
                        string colAValue = rowNode.SelectSingleNode("ColumnA")?.InnerText ?? string.Empty;
                        worksheet.Cells[currentRow, 0].PutValue(colAValue);

                        // Column B - list value
                        string colBValue = rowNode.SelectSingleNode("ColumnB")?.InnerText ?? string.Empty;
                        worksheet.Cells[currentRow, 1].PutValue(colBValue);

                        // Column C - custom string format
                        string colCValue = rowNode.SelectSingleNode("ColumnC")?.InnerText ?? string.Empty;
                        worksheet.Cells[currentRow, 2].PutValue(colCValue);

                        currentRow++;
                    }
                }

                int lastDataRow = currentRow - 1; // Index of the last data row

                // ---------- Validation for Column A (Integer between 1 and 1000) ----------
                CellArea areaA = new CellArea
                {
                    StartRow = startRowIndex,
                    StartColumn = 0,
                    EndRow = lastDataRow,
                    EndColumn = 0
                };
                int validationIndexA = worksheet.Validations.Add(areaA);
                Validation validationA = worksheet.Validations[validationIndexA];
                validationA.Type = Aspose.Cells.ValidationType.WholeNumber;
                validationA.Operator = OperatorType.Between;
                validationA.Formula1 = "1";
                validationA.Formula2 = "1000";
                validationA.InputTitle = "Enter Integer";
                validationA.InputMessage = "Please enter a whole number between 1 and 1000.";
                validationA.ErrorTitle = "Invalid Input";
                validationA.ErrorMessage = "The value must be an integer between 1 and 1000.";
                validationA.ShowError = true;

                // ---------- Validation for Column B (List of predefined options) ----------
                CellArea areaB = new CellArea
                {
                    StartRow = startRowIndex,
                    StartColumn = 1,
                    EndRow = lastDataRow,
                    EndColumn = 1
                };
                int validationIndexB = worksheet.Validations.Add(areaB);
                Validation validationB = worksheet.Validations[validationIndexB];
                validationB.Type = Aspose.Cells.ValidationType.List;
                validationB.Formula1 = "\"Option1,Option2,Option3\"";
                validationB.InputTitle = "Select Option";
                validationB.InputMessage = "Choose one of the allowed options.";
                validationB.ErrorTitle = "Invalid Selection";
                validationB.ErrorMessage = "Please select a value from the list.";
                validationB.ShowError = true;

                // ---------- Validation for Column C (Exactly 5 numeric characters) ----------
                CellArea areaC = new CellArea
                {
                    StartRow = startRowIndex,
                    StartColumn = 2,
                    EndRow = lastDataRow,
                    EndColumn = 2
                };
                int validationIndexC = worksheet.Validations.Add(areaC);
                Validation validationC = worksheet.Validations[validationIndexC];
                validationC.Type = Aspose.Cells.ValidationType.Custom;
                validationC.Formula1 = "=AND(LEN(C2)=5, ISNUMBER(VALUE(C2)))";
                validationC.InputTitle = "Enter Code";
                validationC.InputMessage = "Enter a 5‑digit numeric code.";
                validationC.ErrorTitle = "Invalid Code";
                validationC.ErrorMessage = "The code must be exactly 5 numeric characters.";
                validationC.ShowError = true;

                // Save the workbook
                string outputPath = "ValidatedData.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
