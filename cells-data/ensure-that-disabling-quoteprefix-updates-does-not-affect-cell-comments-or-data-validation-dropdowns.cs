using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 1. Add a comment to cell B2
                // -------------------------------------------------
                // Aspose.Cells manages comments via the Worksheet.Comments collection.
                int commentIndex = sheet.Comments.Add("B2");
                Comment comment = sheet.Comments[commentIndex];
                comment.Author = "Author";
                comment.Note = "This is a comment.";
                comment.Width = 200;
                comment.Height = 100;

                // -------------------------------------------------
                // 2. Add data validation with an in‑cell dropdown to cell C3
                // -------------------------------------------------
                Validation validation = cells["C3"].GetValidation();
                validation.Type = ValidationType.List;
                validation.Formula1 = "Option1,Option2,Option3";
                validation.InCellDropDown = true; // enable dropdown

                // -------------------------------------------------
                // 3. Disable QuotePrefixToStyle
                // -------------------------------------------------
                // When false, a leading single quote is treated as a formatting flag,
                // not as part of the cell's literal value.
                workbook.Settings.QuotePrefixToStyle = false;

                // -------------------------------------------------
                // 4. Put a value that starts with a single quote into cell A1
                // -------------------------------------------------
                cells["A1"].PutValue("'12345"); // leading quote should be removed from value,
                                                // QuotePrefix flag will be set to true.

                // -------------------------------------------------
                // 5. Verify that the comment and validation are unaffected
                // -------------------------------------------------
                Console.WriteLine("Cell A1 Value: " + cells["A1"].StringValue);
                Console.WriteLine("Cell A1 QuotePrefix Style: " + cells["A1"].GetStyle().QuotePrefix);
                Console.WriteLine("Comment in B2: " + sheet.Comments["B2"].Note);
                Console.WriteLine("Validation InCellDropDown for C3: " + validation.InCellDropDown);
                Console.WriteLine("Validation List Formula: " + validation.Formula1);

                // Save the workbook (lifecycle save)
                string outputPath = "QuotePrefix_NoImpact.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}