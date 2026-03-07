using System;
using Aspose.Cells;

namespace AsposeCellsHeaderFooterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a header script containing several commands:
            // &[Page] - current page number
            // &[Pages] - total page count
            // &[Date] - current date
            // &[Time] - current time
            // &"Arial"&12Sample Text - custom text with font settings
            string headerScript = "&[Page] of &[Pages] - &[Date] - &[Time] - &\"Arial\"&12Sample Text";

            // Set the header for the left section (index 0)
            worksheet.PageSetup.SetHeader(0, headerScript);

            // Retrieve the header script back (optional, shows how to get it)
            string retrievedHeader = worksheet.PageSetup.GetHeader(0);

            // Parse the header script into individual commands
            HeaderFooterCommand[] commands = worksheet.PageSetup.GetCommands(retrievedHeader);

            // Display each command's type and associated text (if any)
            Console.WriteLine("Header Commands:");
            foreach (HeaderFooterCommand cmd in commands)
            {
                Console.WriteLine($"- Type: {cmd.Type}");

                // Text commands contain the literal string; other types have empty Text
                if (cmd.Type == HeaderFooterCommandType.Text)
                {
                    Console.WriteLine($"  Text: \"{cmd.Text}\"");
                }
            }

            // Example of setting a footer with a picture command (type = Picture)
            // The picture command uses the syntax &[Picture]
            string footerScript = "&[Picture]";
            worksheet.PageSetup.SetFooter(0, footerScript);

            // Parse and display footer commands
            HeaderFooterCommand[] footerCommands = worksheet.PageSetup.GetCommands(worksheet.PageSetup.GetFooter(0));
            Console.WriteLine("\nFooter Commands:");
            foreach (HeaderFooterCommand cmd in footerCommands)
            {
                Console.WriteLine($"- Type: {cmd.Type}");
                if (cmd.Type == HeaderFooterCommandType.Text)
                {
                    Console.WriteLine($"  Text: \"{cmd.Text}\"");
                }
            }

            // Save the workbook to verify that header/footer are applied
            workbook.Save("HeaderFooterDemo.xlsx");
        }
    }
}