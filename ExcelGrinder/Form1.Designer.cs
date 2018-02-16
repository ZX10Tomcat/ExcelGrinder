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
            this.InfoLabel = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ExcelOutputView = new System.Windows.Forms.DataGridView();
            this.Grindbtn = new System.Windows.Forms.Button();
            this.ChoseFolderbtn = new System.Windows.Forms.Button();
            this.Cancelbtn = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExcelRuleBookView)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExcelOutputView)).BeginInit();
            this.SuspendLayout();
            // 
            // ChoseFilebtn
            // 
            this.ChoseFilebtn.Location = new System.Drawing.Point(9, 29);
            this.ChoseFilebtn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ChoseFilebtn.Name = "ChoseFilebtn";
            this.ChoseFilebtn.Size = new System.Drawing.Size(273, 71);
            this.ChoseFilebtn.TabIndex = 0;
            this.ChoseFilebtn.Text = "Открыть файл с фамилиями";
            this.ChoseFilebtn.UseVisualStyleBackColor = true;
            this.ChoseFilebtn.Click += new System.EventHandler(this.ChoseFilebtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ExcelRuleBookView);
            this.groupBox1.Controls.Add(this.ChoseFilebtn);
            this.groupBox1.Location = new System.Drawing.Point(20, 20);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(486, 686);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Список фамилий";
            // 
            // ExcelRuleBookView
            // 
            this.ExcelRuleBookView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ExcelRuleBookView.Location = new System.Drawing.Point(10, 111);
            this.ExcelRuleBookView.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ExcelRuleBookView.Name = "ExcelRuleBookView";
            this.ExcelRuleBookView.Size = new System.Drawing.Size(466, 562);
            this.ExcelRuleBookView.TabIndex = 1;
            // 
            // InfoLabel
            // 
            this.InfoLabel.AutoSize = true;
            this.InfoLabel.Location = new System.Drawing.Point(20, 722);
            this.InfoLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.InfoLabel.Name = "InfoLabel";
            this.InfoLabel.Size = new System.Drawing.Size(0, 20);
            this.InfoLabel.TabIndex = 2;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ExcelOutputView);
            this.groupBox2.Controls.Add(this.Grindbtn);
            this.groupBox2.Controls.Add(this.ChoseFolderbtn);
            this.groupBox2.Location = new System.Drawing.Point(516, 20);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Size = new System.Drawing.Size(788, 686);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Вывод";
            // 
            // ExcelOutputView
            // 
            this.ExcelOutputView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ExcelOutputView.Location = new System.Drawing.Point(9, 108);
            this.ExcelOutputView.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ExcelOutputView.Name = "ExcelOutputView";
            this.ExcelOutputView.Size = new System.Drawing.Size(770, 562);
            this.ExcelOutputView.TabIndex = 2;
            // 
            // Grindbtn
            // 
            this.Grindbtn.BackColor = System.Drawing.Color.Lavender;
            this.Grindbtn.Location = new System.Drawing.Point(321, 29);
            this.Grindbtn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Grindbtn.Name = "Grindbtn";
            this.Grindbtn.Size = new System.Drawing.Size(458, 69);
            this.Grindbtn.TabIndex = 1;
            this.Grindbtn.Text = "НАЧАТЬ!";
            this.Grindbtn.UseVisualStyleBackColor = false;
            this.Grindbtn.Click += new System.EventHandler(this.Grindbtn_Click);
            // 
            // ChoseFolderbtn
            // 
            this.ChoseFolderbtn.Location = new System.Drawing.Point(9, 31);
            this.ChoseFolderbtn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ChoseFolderbtn.Name = "ChoseFolderbtn";
            this.ChoseFolderbtn.Size = new System.Drawing.Size(272, 69);
            this.ChoseFolderbtn.TabIndex = 0;
            this.ChoseFolderbtn.Text = "Выбрать папку с файлами";
            this.ChoseFolderbtn.UseVisualStyleBackColor = true;
            this.ChoseFolderbtn.Click += new System.EventHandler(this.ChoseFolderbtn_Click);
            // 
            // Cancelbtn
            // 
            this.Cancelbtn.BackColor = System.Drawing.Color.Red;
            this.Cancelbtn.Location = new System.Drawing.Point(1150, 715);
            this.Cancelbtn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Cancelbtn.Name = "Cancelbtn";
            this.Cancelbtn.Size = new System.Drawing.Size(153, 40);
            this.Cancelbtn.TabIndex = 4;
            this.Cancelbtn.Text = "ОТМЕНА";
            this.Cancelbtn.UseVisualStyleBackColor = false;
            this.Cancelbtn.Click += new System.EventHandler(this.Cancelbtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1322, 760);
            this.Controls.Add(this.Cancelbtn);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.InfoLabel);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel grinder";
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExcelRuleBookView)).EndInit();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExcelOutputView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ChoseFilebtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView ExcelRuleBookView;
        private System.Windows.Forms.Label InfoLabel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button Grindbtn;
        private System.Windows.Forms.Button ChoseFolderbtn;
        private System.Windows.Forms.Button Cancelbtn;
        private System.Windows.Forms.DataGridView ExcelOutputView;
    }
}

