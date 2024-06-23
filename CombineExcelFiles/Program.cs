using ClosedXML.Excel;
try
{

    string entry;
    while (true)
    {
        Console.Write("Result Excel Sheet Names From ?\n1 => From Excel Files Names\n2 => From Excel Files Sheets Names\n-->");
        entry = Console.ReadLine()!;

        if (entry == "1" || entry == "2") break;

        Console.WriteLine("*** Wrong Entry ***\n");
    }

    var excelFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx");
    if (excelFiles.Length == 0) return;

    var finalWorkbook = new XLWorkbook
    {
        RightToLeft = true
    };

    var percentSteps = 100 / excelFiles.Length;

    for (int i = 0; i < excelFiles.Length; i++)
    {
        var workbook = new XLWorkbook(excelFiles[i]);

        foreach (var worksheet in workbook.Worksheets)
        {
            if (entry == "1")
            {
                worksheet.CopyTo(finalWorkbook, excelFiles[i].Split('\\').Last().Replace(".xlsx", ""));
            }
            else
            {
                worksheet.CopyTo(finalWorkbook, worksheet.Name);
            }
        }
        Console.WriteLine($"Progress : {i + 1 * percentSteps}%");
    }

    finalWorkbook.SaveAs("Result/Result.xlsx");
    Console.WriteLine("*** Operation Successful ***");
}
catch (Exception e)
{
    File.WriteAllText($"{Directory.GetCurrentDirectory()}/Log.txt", e.Message);
    Console.WriteLine("*** Error ! See Log.txt File ***\nPress any key to close");
    Console.ReadKey();
}