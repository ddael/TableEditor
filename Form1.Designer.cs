namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            ImportButton = new Button();
            ExportButton = new Button();
            filteredGrid = new DataGridView();
            label1 = new Label();
            checkBox1 = new CheckBox();
            checkBox2 = new CheckBox();
            checkBox3 = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)filteredGrid).BeginInit();
            SuspendLayout();
            // 
            // ImportButton
            // 
            ImportButton.Location = new Point(45, 358);
            ImportButton.Name = "ImportButton";
            ImportButton.Size = new Size(105, 40);
            ImportButton.TabIndex = 0;
            ImportButton.Text = "Импорт";
            ImportButton.UseVisualStyleBackColor = true;
            ImportButton.Click += ImportButton_Click;
            // 
            // ExportButton
            // 
            ExportButton.Location = new Point(645, 358);
            ExportButton.Name = "ExportButton";
            ExportButton.Size = new Size(105, 40);
            ExportButton.TabIndex = 1;
            ExportButton.Text = "Экспорт";
            ExportButton.UseVisualStyleBackColor = true;
            ExportButton.Click += ExportButton_Click;
            // 
            // filteredGrid
            // 
            filteredGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            filteredGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            filteredGrid.Location = new Point(45, 45);
            filteredGrid.Name = "filteredGrid";
            filteredGrid.Size = new Size(705, 284);
            filteredGrid.TabIndex = 3;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(267, 25);
            label1.Name = "label1";
            label1.Size = new Size(254, 15);
            label1.TabIndex = 4;
            label1.Text = "Форматированная таблица(Work in progress)";
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(299, 365);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(189, 19);
            checkBox1.TabIndex = 5;
            checkBox1.Text = "Режим разбивки на кластеры";
            checkBox1.UseVisualStyleBackColor = true;
            checkBox1.CheckedChanged += checkBox1_CheckedChanged;
            // 
            // checkBox2
            // 
            checkBox2.AutoSize = true;
            checkBox2.Location = new Point(299, 390);
            checkBox2.Name = "checkBox2";
            checkBox2.Size = new Size(175, 19);
            checkBox2.TabIndex = 6;
            checkBox2.Text = "Регионы вместо кластеров";
            checkBox2.UseVisualStyleBackColor = true;
            checkBox2.CheckedChanged += checkBox2_CheckedChanged;
            // 
            // checkBox3
            // 
            checkBox3.AutoSize = true;
            checkBox3.Location = new Point(299, 340);
            checkBox3.Name = "checkBox3";
            checkBox3.Size = new Size(127, 19);
            checkBox3.TabIndex = 7;
            checkBox3.Text = "Режим аттестации";
            checkBox3.UseVisualStyleBackColor = true;
            checkBox3.CheckedChanged += checkBox3_CheckedChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(784, 421);
            Controls.Add(checkBox3);
            Controls.Add(checkBox2);
            Controls.Add(checkBox1);
            Controls.Add(label1);
            Controls.Add(filteredGrid);
            Controls.Add(ExportButton);
            Controls.Add(ImportButton);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            Text = "Вычищатель таблиц";
            ((System.ComponentModel.ISupportInitialize)filteredGrid).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button ImportButton;
        private Button ExportButton;
        private DataGridView filteredGrid;
        private Label label1;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private CheckBox checkBox3;
    }
}