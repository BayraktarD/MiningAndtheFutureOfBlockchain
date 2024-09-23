using ClosedXML.Excel;

string filePath = "C:/Users/dorin.bayraktar/Downloads/CTECH203_Activity_Solve+a+Computational+Puzzle.xlsx"; // Replace with your actual file path
using (var workbook = new XLWorkbook(filePath))
{
    var worksheet = workbook.Worksheet(1); // Instructions Worksheet

    // Starting and ending rows for x values and hash checks
    int startRow = 28;
    int endRow = 56;

    // Start x value with the first 10-digit number
    long currentX = 10000000000;

    // Loop to find and set x values until all hashes are valid
    for (int i = startRow; i <= endRow; i++)
    {
        long hashValue = -1;

        // Continue incrementing x until a valid hash is found
        while (hashValue % 100 != 0)
        {
            // Write the current x value to column H (8th column)
            worksheet.Cell(i, 8).Value = currentX.ToString(); // Write x as string to maintain the formatting
            Console.WriteLine($"Row {i}: Writing x = {currentX} in column H");

            // Save the workbook to allow Excel formulas to recalculate
            workbook.Save();

            // Read the recalculated hash value from column I (9th column)
            string hashString = worksheet.Cell(i, 9).GetValue<string>();

            // Parse the hash value
            if (long.TryParse(hashString, out hashValue))
            {
                if (hashValue % 100 == 0)
                {
                    Console.WriteLine($"Row {i}: Valid x = {currentX}, Hash = {hashValue} (Ends with two+ zeros)");
                    break;
                }
            }
            else
            {
                Console.WriteLine($"Row {i}: Unable to parse hash value. Read: {hashString}");
            }

            // Increment x by 1
            currentX++;
        }

        // Move to the next row with currentX + 1
        currentX++;
    }

    // Save the final changes to the workbook
    workbook.Save();
    Console.WriteLine("\nAll rows have valid hashes ending with two or more zeros!");
}

Console.WriteLine("Done!");
