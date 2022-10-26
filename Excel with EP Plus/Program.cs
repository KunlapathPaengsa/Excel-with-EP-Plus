using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using static System.Net.Mime.MediaTypeNames;

class Program
{
    static void Main(string[] args)
    {
        var Articles = new[]
        {
                new {
                    Id = "101", Name = "C++"
                },
                new {
                    Id = "102", Name = "Python"
                },
                new {
                    Id = "103", Name = "Java Script"
                },
                new {
                    Id = "104", Name = "GO"
                },
                new {
                    Id = "105", Name = "Java"
                },
                new {
                    Id = "106", Name = "C#"
                }
            };

        // Creating an instance
        // of ExcelPackage
        ExcelPackage excel = new ExcelPackage();

        // name of the sheet
        var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

        // setting the properties
        // of the work sheet
        workSheet.TabColor = Color.Black;
        workSheet.DefaultRowHeight = 12;

        // Setting the properties
        // of the first row
        workSheet.Row(1).Height = 20;
        workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        workSheet.Row(1).Style.Font.Bold = true;
        workSheet.Row(1).Style.Font.Color.SetColor(Color.White);
        workSheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
        //workSheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        using (ExcelRange Rng = workSheet.Cells[5, 2, 8, 4])
        {
            //Rng.Value = "Text Color & Background Color";
            //Rng.Merge = true;
            //Rng.Style.Font.Bold = true;
            Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
            Rng.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        }


        // Header of the Excel sheet 
        workSheet.Cells[1, 1].Value = "S.No";
        workSheet.Cells[1, 2].Value = "Id";
        workSheet.Cells[1, 3].Value = "Name";

        // Inserting the article data into excel
        // sheet by using the for each loop
        // As we have values to the first row
        // we will start with second row
        int recordIndex = 2;
        //int recordIndex = 1;

        foreach (var article in Articles)
        {
            workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
            workSheet.Cells[recordIndex, 2].Value = article.Id;
            workSheet.Cells[recordIndex, 3].Value = article.Name;
            recordIndex++;
        }

        // By default, the column width is not
        // set to auto fit for the content
        // of the range, so we are using
        // AutoFit() method here.
        workSheet.Column(1).AutoFit();
        workSheet.Column(2).AutoFit();
        workSheet.Column(3).AutoFit();

        // file name with .xlsx extension
        string p_strPath = "D:\\POC\\Excel with EP Plus\\Excel with EP Plus\\Models\\geeksforgeeks9.xlsx";

        if (File.Exists(p_strPath))
            File.Delete(p_strPath);

        // Create excel file on physical disk
        FileStream objFileStrm = File.Create(p_strPath);
        objFileStrm.Close();

        // Write content to excel file
        File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
        //Close Excel package
        excel.Dispose();
        Console.WriteLine("Generate file SUCCESS!");
        ReadFile(p_strPath);
    }

    private static void ReadFile(string path)
    {
        //StringBuilder builderPath = new StringBuilder().Append("@\"").Append(path).Append("\"") ;
        FileInfo excel = new FileInfo(@"D:\POC\Excel with EP Plus\Excel with EP Plus\Models\geeksforgeeks9.xlsx");//builderPath.ToString());
        using (var excelPack = new ExcelPackage()) //ExcelPackage package = new ExcelPackage(existingFile))
        {
            using (var package = new ExcelPackage(excel))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 1; row < rowCount; row++)
                {
                    for (int col = 1; col < colCount; col++)
                    {
                        Console.Write($"| {worksheet.Cells[row, col].Value?.ToString()}");
                    }
                    Console.WriteLine();
                }
            }
        }
    }
}

