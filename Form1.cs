namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private List<Student> data = [];
        private List<Student> filtered = [];

        private string Name =string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" };
            Name = dialog.FileName;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var processor = new Processor();
                data = Processor.ReadFromFile(dialog.FileName);

                filtered = data.Select(s => new Student
                {
                    FullName = s.FullName,
                    Group = s.Group,
                    Cohort = s.Cohort,
                    Region = s.Region,
                    Results = s.Results
                }
                ).ToList();

                filteredGrid.DataSource = null;
                filteredGrid.DataSource = filtered;
                
                MessageBox.Show($"Импортировано записей: {data.Count}");
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            if (data.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.");
                return;
            }

            using var dialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx", FileName = Name + "_formated.xlsx" };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _ = new Processor();
                Processor.WriteToFile(dialog.FileName, data);
                MessageBox.Show("Файл сохранён.");
            }
        }

    }
}
