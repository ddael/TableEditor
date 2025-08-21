using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace WinFormsApp1
{
    public class Processor
    {
        public static List<Student> ReadFromFile(string filePath)
        {
            var regFN = new Regex(@"ФИО");
            var indexFN = 0;

            var regTag = new Regex(@"Теги");
            var indexTag = 0;

            var regRegion = new Regex(@"Регион");
            var indexRegion = 0;

            var regCohort = new Regex(@"Поток");
            var indexCohort = 0;

            var regWS = new Regex(@".*\(Значение\)$");
            var indexWS = new Dictionary<int, string>();

            var result = new List<Student>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var headerRow = worksheet.Row(1); // Colomns title row

            foreach (var cell in headerRow.CellsUsed())
            {
                if (regFN.IsMatch(cell.GetString()))
                {
                    indexFN = cell.Address.ColumnNumber;
                }

                if (regTag.IsMatch(cell.GetString()))
                {
                    indexTag = cell.Address.ColumnNumber;
                }

                if (regRegion.IsMatch(cell.GetString()))
                {
                    indexRegion = cell.Address.ColumnNumber;
                }

                if (regCohort.IsMatch(cell.GetString()))
                {
                    indexCohort = cell.Address.ColumnNumber;
                }

                if (regWS.IsMatch(cell.GetString()))
                {
                    indexWS.Add(cell.Address.ColumnNumber, cell.GetValue<string>());
                }
            }

            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

            foreach (var row in rows)
            {
                var fullName = row.Cell(indexFN).GetString();
                var group = row.Cell(indexTag).GetString();
                group = Regex.Match(group, @"[а-яА-ЯёЁ]{2,4}/\d{2}").ToString();
                var region = row.Cell(indexRegion).GetString();
                var cohort = row.Cell(indexCohort).GetString();
                if (cohort == "0.00")
                {
                    continue;
                }

                var results = new Dictionary<string, string>();

                foreach (var cell in row.CellsUsed())
                {
                    foreach (var item in indexWS.Keys)
                    {
                        if (cell.Address.ColumnNumber == item)
                        {
                            results.Add(indexWS[item], cell.GetString());
                        }
                    }
                }

                result.Add(new Student
                {
                    FullName = fullName,
                    Group = group,
                    Region = region,
                    Cohort = cohort,
                    Results = ConvertDictionary(results)
                });
            }

            return result;
        }

        public static void WriteToFile(string filePath, List<Student> Students)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Filtered");

            // Write header
            worksheet.Cell(1, 1).Value = "ФИО";
            worksheet.Cell(1, 2).Value = "Группа";
            worksheet.Cell(1, 3).Value = "Поток";
            worksheet.Cell(1, 4).Value = "Регион";
            worksheet.Cell(1, 5).Value = "Кол-во\nсданных\nработ";
            List<string> Headers = Students[1].Results.Keys.ToList();

            for (int i = 0; i < Headers.Count; i++)
            {
                worksheet.Cell(1, 5 + (i + 1)).Value = Headers[i];
            }

            // Write data
            for (int i = 0; i < Students.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = Students[i].FullName;
                worksheet.Cell(i + 2, 2).Value = Students[i].Group;
                worksheet.Cell(i + 2, 3).Value = Students[i].Cohort;
                worksheet.Cell(i + 2, 4).Value = Students[i].Region;
                worksheet.Cell(i + 2, 5).Value = Students[i].GetRatio();
                List<WorkStatus> StudStat = Students[i].Results.Values.ToList();
                for (int j = 0; j < StudStat.Count; j++)
                {
                    var cell = worksheet.Cell(i + 2, 5 + (j + 1));
                    cell.Value = StudStat[j].ToString().Replace("_", " ");

                    if (StudStat[j] == WorkStatus.Зачтено)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C6EFCE");
                    }
                    else if (StudStat[j] == WorkStatus.На_проверке)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFEB9D");
                    }
                    else
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC7CE");
                    }
                }
            }

            var table = worksheet.Range(1, 1, Students.Count + 1, Headers.Count + 5).AsTable();
            table.Name = "Учащиеся";
            table.ShowAutoFilter = true;
            table.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            table.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

            worksheet.Columns(1, 5).AdjustToContents();

            worksheet.Row(1).AdjustToContents();
            worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Row(1).Style.Font.Bold = true; 

            worksheet.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Column(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int lastColumn = 5 + (Students.Count > 0 ? Students[0].Results.Count : 0);
            worksheet.Columns(6, lastColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Columns(6, lastColumn).Width = 20;

            workbook.SaveAs(filePath);
        }

        private static WorkStatus ConvertToWorkStatus(string status)
        {
            return status switch
            {
                "Зачтено" => WorkStatus.Зачтено,
                "Не зачтено" or "Не_зачтено" => WorkStatus.Не_зачтено,
                "Не сдано" or "Не_сдано" => WorkStatus.Не_сдано,
                "На проверке" or "На_проверке" => WorkStatus.На_проверке,
                "Не доступно" or "Не_доступно" => WorkStatus.Не_доступно,
                _ => throw new ArgumentException($"Неизвестный статус: {status}")
            };
        }

        private static Dictionary<string, WorkStatus> ConvertDictionary(Dictionary<string, string> stringDict)
        {
            return stringDict.ToDictionary(
                kvp => kvp.Key,
                kvp => ConvertToWorkStatus(kvp.Value)
                );
        }
    }
}
