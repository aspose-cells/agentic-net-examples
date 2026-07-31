// Title: Aspose.Cells .NET: Create and Rename a Named Range (A1:C10) to ReportData
// Description: Demonstrates how to create a new Workbook, define a range covering cells A1:C10, assign it the name "DataSet", retrieve the corresponding Name object, rename it to "ReportData" using the Name.Text property, and save the file as NamedRangeDemo.xlsx.
// Keywords: Aspose.Cells named range C# | rename named range Aspose.Cells | Name.Text property | create range A1:C10 | Aspose.Cells workbook save | C# Excel automation | named range example .NET
// Common Searches: Aspose.Cells create named range A1:C10 | How to rename a named range with Name.Text in Aspose.Cells | C# code to change named range name in Excel file | Aspose.Cells example for managing named ranges
// Developer Intent: Programmatically create a named range for A1:C10 and change its identifier to ReportData.
// Use Cases: Prepare a template workbook with a placeholder range that can be renamed during a reporting workflow. | Allow dynamic renaming of data blocks referenced by formulas without altering cell references. | Support user‑driven range naming where an initial name is set and later updated based on business logic.
// AI Prompts: Generate C# code using Aspose.Cells to define a range A1:C10, name it DataSet, then rename it to ReportData via Name.Text. | Explain the steps to retrieve a Name object from a workbook's Names collection and modify its Text property in Aspose.Cells .NET. | Provide a concise tutorial for creating, renaming, and saving a workbook with a named range using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDemo
{
    // Demonstrates how to create a new Workbook, define a range covering cells A1:C10, assign it the name "DataSet", retrieve the corresponding Name object, rename it to "ReportData" using the Name.Text property, and save the file as NamedRangeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range that covers cells A1:C10
                // Parameters: first row index, first column index, total rows, total columns
                Aspose.Cells.Range dataSetRange = worksheet.Cells.CreateRange(0, 0, 10, 3);

                // Assign an initial name to the range
                dataSetRange.Name = "DataSet";

                // Retrieve the Name object that was created for the range
                Name nameObj = workbook.Worksheets.Names["DataSet"];

                // Rename the range by setting the Text property of the Name object
                if (nameObj != null)
                {
                    nameObj.Text = "ReportData";
                }

                // Save the workbook to a file
                string outputPath = "NamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
