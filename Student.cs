using System.ComponentModel.DataAnnotations;

namespace WinFormsApp1
{
    public class Student
    {
        [Required]
        public string FullName { get; set; } = string.Empty;
        public string Cluster { get; set; } = string.Empty;
        public int Group { get; set; }
        public string Region { get; set; } = string.Empty;
        public string Cohort { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public Results Results { get; set; } = [];

        public int GetRatio()
        {
            return Results.Ratio;
        }

        public string? GetFullGroup(char splitter)
        {
            if (string.IsNullOrEmpty(Tag) || Group == 0)
            {
                return null;
            }
            return $"{Tag}{splitter}{Group:D2}";
        }

        public Results FilterKeys(string pattern)
        {
            return Results.Filter(pattern);
        }

    }

    public enum WorkStatus
    {
        Зачтено,
        Не_зачтено,
        Не_сдано,
        На_проверке,
        Недоступно
    }
}
