// Title: How to update or add an integer custom property (ProjectId) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update existing custom document property value in Excel | Add integer custom property to Excel file with Aspose.Cells .NET | Set ProjectId custom property programmatically in a workbook using Aspose.Cells | Check and create missing custom document property in Excel via C# Aspose.Cells | Modify custom document properties of .xlsx using Aspose.Cells API
// Tags: update custom document property Aspose.Cells | add integer custom property Excel .NET | custom property index lookup Aspose.Cells | modify workbook custom properties C# | Aspose.Cells custom property handling

using Aspose.Cells;
using System;

// The example loads an Excel workbook, accesses its CustomDocumentProperties collection, searches for a property named "ProjectId", updates its integer value to 12345 if it exists, or adds the property with that value if it does not, and then saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of custom document properties
        var customProps = workbook.CustomDocumentProperties;

        // Try to find the index of the property named "ProjectId"
        int propIndex = customProps.IndexOf("ProjectId");

        if (propIndex != -1)
        {
            // Property exists – update its integer value
            customProps[propIndex].Value = 12345; // set the new integer value here
        }
        else
        {
            // Property does not exist – optionally add it with the desired value
            customProps.Add("ProjectId", 12345);
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
