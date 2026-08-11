// Title: Enumerate all worksheets in an Aspose.Cells workbook and log each Name and TabId (C#)
// Description: This C# example creates a workbook, assigns custom names and TabId values to worksheets, then iterates through the Worksheets collection, writing each sheet’s Name and TabId to the console before saving the file as AuditWorkbook.xlsx.
// Keywords: Aspose.Cells | C# | worksheet TabId | list worksheet names | audit workbook | enumerate worksheets | Aspose.Cells .NET | Worksheet TabId property | iterate worksheets
// Common Searches: Aspose.Cells get worksheet TabId | How to loop through worksheets in Aspose.Cells C# | Retrieve worksheet names and TabIds Aspose.Cells | Audit worksheet identifiers Aspose.Cells | Save workbook after enumerating worksheets
// Developer Intent: The developer needs to enumerate every worksheet in a workbook and capture its Name and TabId for auditing, validation, or reporting purposes.
// Use Cases: Create an audit log of worksheet identifiers before distribution | Verify worksheet presence and correct TabId in automated tests | Generate compliance documentation of workbook structure | Export worksheet metadata to external systems
// AI Prompts: Generate C# code that writes worksheet names and TabIds to a CSV using Aspose.Cells | Show how to filter worksheets by a range of TabId values while iterating | Explain best practices for handling missing TabId or access errors in Aspose.Cells | Provide example of logging worksheet metadata to a file instead of the console

using System;
using Aspose.Cells;

// This C# example creates a workbook, assigns custom names and TabId values to worksheets, then iterates through the Worksheets collection, writing each sheet’s Name and TabId to the console before saving the file as AuditWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Example setup: assign names and TabIds to worksheets
            Worksheet first = workbook.Worksheets[0];
            first.Name = "FirstSheet";
            first.TabId = 101;

            // Add a second worksheet and set its TabId
            Worksheet second = workbook.Worksheets.Add("SecondSheet"); // Add returns the new worksheet
            second.TabId = 102;

            // Iterate through all worksheets and log their Name and TabId
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet Name: {sheet.Name}, TabId: {sheet.TabId}");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("AuditWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
