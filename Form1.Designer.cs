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
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(784, 421);
            Controls.Add(label1);
            Controls.Add(filteredGrid);
            Controls.Add(ExportButton);
            Controls.Add(ImportButton);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            Text = "Вычищатель таблиц 1.0";
            ((System.ComponentModel.ISupportInitialize)filteredGrid).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button ImportButton;
        private Button ExportButton;
        private DataGridView filteredGrid;
        private Label label1;
    }
}