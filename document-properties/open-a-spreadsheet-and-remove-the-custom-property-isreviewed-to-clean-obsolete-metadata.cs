// Title: Delete the 'IsReviewed' custom document property from an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# code to remove a custom document property from an .xlsx file | how to programmatically delete the IsReviewed property in an Excel workbook using Aspose | remove obsolete custom properties from Excel using Aspose.Cells .NET API | C# example for clearing specific custom document properties in a spreadsheet | Aspose.Cells remove custom property if it exists before saving workbook
// Tags: Aspose.Cells delete custom document property C# | remove specific Excel custom property programmatically | custom document properties API Aspose.Cells | clean obsolete spreadsheet metadata .NET | Workbook custom property management Aspose | C# remove IsReviewed property from workbook

using Aspose.Cells;
using System;

// The sample loads an Excel file (input.xlsx) into an Aspose.Cells Workbook, checks the CustomDocumentProperties collection for a property named "IsReviewed", removes it when present, and saves the cleaned workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of custom document properties
        var customProps = workbook.CustomDocumentProperties;

        // Remove the custom property named "IsReviewed" if it exists
        if (customProps.Contains("IsReviewed"))
        {
            customProps.Remove("IsReviewed");
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
