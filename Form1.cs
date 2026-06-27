using DocumentFormat.OpenXml.Drawing;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private bool MultiFileMode { get; set; } = false;

        private List<Student> data = [];
        private List<Student> filtered = [];

        private string Name = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            if (!MultiFileMode)
            {
                using var dialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(dialog.SafeFileName);
                    data = Processor.ReadFromFile(dialog.FileName);

                    filtered = data.Select(s => new Student
                    {
                        FullName = s.FullName,
                        Cluster = s.Cluster,
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
            else
            {
                //using var dialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls" };
                //if (dialog.ShowDialog() == DialogResult.OK)
                //{
                //    data.AddRange(Processor.ReadFromFile(Name));
                //}
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            if(!MultiFileMode)
            {
                if (data.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта.");
                    return;
                }

                using var dialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx", FileName = $"{Name}_formated" };
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _ = new Processor();
                    Processor.WriteToFile(dialog.FileName, data);
                    MessageBox.Show("Файл сохранён.");
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            MultiFileMode = checkBox1.Checked;
        }
    }
}
