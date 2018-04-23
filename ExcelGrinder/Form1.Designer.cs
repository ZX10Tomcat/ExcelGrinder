namespace ExcelGrinder
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
            this.ChoseFilebtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ExcelRuleBookView = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Grindbtn = new System.Windows.Forms.Button();
            this.ChoseFolderbtn = new System.Windows.Forms.Button();
            this.Cancelbtn = new System.Windows.Forms.Button();
            this.testExcelbtn = new System.Windows.Forms.Button();
            this.InfoText = new System.Windows.Forms.RichTextBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExcelRuleBookView)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ChoseFilebtn
            // 
            this.ChoseFilebtn.Location = new System.Drawing.Point(6, 19);
            this.ChoseFilebtn.Name = "ChoseFilebtn";
            this.ChoseFilebtn.Size = new System.Drawing.Size(182, 46);
            this.ChoseFilebtn.TabIndex = 0;
            this.ChoseFilebtn.Text = "Открыть файл с фамилиями";
            this.ChoseFilebtn.UseVisualStyleBackColor = true;
            this.ChoseFilebtn.Click += new System.EventHandler(this.ChoseFilebtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ExcelRuleBookView);
            this.groupBox1.Controls.Add(this.ChoseFilebtn);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(324, 446);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Список фамилий";
            // 
            // ExcelRuleBookView
            // 
            this.ExcelRuleBookView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ExcelRuleBookView.Location = new System.Drawing.Point(7, 72);
            this.ExcelRuleBookView.Name = "ExcelRuleBookView";
            this.ExcelRuleBookView.Size = new System.Drawing.Size(311, 365);
            this.ExcelRuleBookView.TabIndex = 1;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.InfoText);
            this.groupBox2.Controls.Add(this.Grindbtn);
            this.groupBox2.Controls.Add(this.ChoseFolderbtn);
            this.groupBox2.Location = new System.Drawing.Point(344, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(525, 446);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Вывод";
            // 
            // Grindbtn
            // 
            this.Grindbtn.BackColor = System.Drawing.Color.Lavender;
            this.Grindbtn.Location = new System.Drawing.Point(214, 19);
            this.Grindbtn.Name = "Grindbtn";
            this.Grindbtn.Size = new System.Drawing.Size(305, 45);
            this.Grindbtn.TabIndex = 1;
            this.Grindbtn.Text = "НАЧАТЬ!";
            this.Grindbtn.UseVisualStyleBackColor = false;
            this.Grindbtn.Click += new System.EventHandler(this.Grindbtn_ClickAsync);
            // 
            // ChoseFolderbtn
            // 
            this.ChoseFolderbtn.Location = new System.Drawing.Point(6, 20);
            this.ChoseFolderbtn.Name = "ChoseFolderbtn";
            this.ChoseFolderbtn.Size = new System.Drawing.Size(181, 45);
            this.ChoseFolderbtn.TabIndex = 0;
            this.ChoseFolderbtn.Text = "Выбрать папку с файлами";
            this.ChoseFolderbtn.UseVisualStyleBackColor = true;
            this.ChoseFolderbtn.Click += new System.EventHandler(this.ChoseFolderbtn_Click);
            // 
            // Cancelbtn
            // 
            this.Cancelbtn.BackColor = System.Drawing.Color.Red;
            this.Cancelbtn.Location = new System.Drawing.Point(767, 463);
            this.Cancelbtn.Name = "Cancelbtn";
            this.Cancelbtn.Size = new System.Drawing.Size(102, 26);
            this.Cancelbtn.TabIndex = 4;
            this.Cancelbtn.Text = "ОТМЕНА";
            this.Cancelbtn.UseVisualStyleBackColor = false;
            this.Cancelbtn.Click += new System.EventHandler(this.Cancelbtn_Click);
            // 
            // testExcelbtn
            // 
            this.testExcelbtn.Location = new System.Drawing.Point(664, 464);
            this.testExcelbtn.Margin = new System.Windows.Forms.Padding(2);
            this.testExcelbtn.Name = "testExcelbtn";
            this.testExcelbtn.Size = new System.Drawing.Size(98, 26);
            this.testExcelbtn.TabIndex = 3;
            this.testExcelbtn.Text = "Тест Excel";
            this.testExcelbtn.UseVisualStyleBackColor = true;
            this.testExcelbtn.Click += new System.EventHandler(this.TestExcelbtn_Click);
            // 
            // InfoText
            // 
            this.InfoText.Location = new System.Drawing.Point(7, 72);
            this.InfoText.Name = "InfoText";
            this.InfoText.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.InfoText.Size = new System.Drawing.Size(518, 365);
            this.InfoText.TabIndex = 3;
            this.InfoText.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(881, 497);
            this.Controls.Add(this.testExcelbtn);
            this.Controls.Add(this.Cancelbtn);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel grinder";
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExcelRuleBookView)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ChoseFilebtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView ExcelRuleBookView;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button Grindbtn;
        private System.Windows.Forms.Button ChoseFolderbtn;
        private System.Windows.Forms.Button Cancelbtn;
        private System.Windows.Forms.Button testExcelbtn;
        private System.Windows.Forms.RichTextBox InfoText;
    }
}

