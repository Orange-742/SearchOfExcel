
namespace SearchOfExcel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Start = new System.Windows.Forms.Button();
            this.SearchWord = new System.Windows.Forms.TextBox();
            this.labelBox = new System.Windows.Forms.ComboBox();
            this.lab = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.FILENAME = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.Show = new System.Windows.Forms.TextBox();
            this.end = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Start
            // 
            this.Start.Location = new System.Drawing.Point(233, 125);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(75, 23);
            this.Start.TabIndex = 0;
            this.Start.Text = "搜索";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // SearchWord
            // 
            this.SearchWord.Location = new System.Drawing.Point(32, 125);
            this.SearchWord.Name = "SearchWord";
            this.SearchWord.Size = new System.Drawing.Size(181, 21);
            this.SearchWord.TabIndex = 3;
            // 
            // labelBox
            // 
            this.labelBox.FormattingEnabled = true;
            this.labelBox.Location = new System.Drawing.Point(114, 61);
            this.labelBox.Name = "labelBox";
            this.labelBox.Size = new System.Drawing.Size(194, 20);
            this.labelBox.TabIndex = 4;
            // 
            // lab
            // 
            this.lab.AutoSize = true;
            this.lab.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lab.Location = new System.Drawing.Point(29, 62);
            this.lab.Name = "lab";
            this.lab.Size = new System.Drawing.Size(71, 16);
            this.lab.TabIndex = 6;
            this.lab.Text = "选择类别";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(32, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(95, 23);
            this.button1.TabIndex = 18;
            this.button1.Text = "选择excel文件";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FILENAME
            // 
            this.FILENAME.AutoSize = true;
            this.FILENAME.Location = new System.Drawing.Point(133, 28);
            this.FILENAME.Name = "FILENAME";
            this.FILENAME.Size = new System.Drawing.Size(65, 12);
            this.FILENAME.TabIndex = 19;
            this.FILENAME.Text = "未选择文件";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(29, 95);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(119, 16);
            this.label10.TabIndex = 20;
            this.label10.Text = "输入搜索的词：";
            // 
            // Show
            // 
            this.Show.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Show.Location = new System.Drawing.Point(326, 12);
            this.Show.Multiline = true;
            this.Show.Name = "Show";
            this.Show.ReadOnly = true;
            this.Show.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.Show.Size = new System.Drawing.Size(292, 336);
            this.Show.TabIndex = 21;
            // 
            // end
            // 
            this.end.Location = new System.Drawing.Point(233, 166);
            this.end.Name = "end";
            this.end.Size = new System.Drawing.Size(75, 23);
            this.end.TabIndex = 22;
            this.end.Text = "关闭";
            this.end.UseVisualStyleBackColor = true;
            this.end.Click += new System.EventHandler(this.end_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(630, 360);
            this.ControlBox = false;
            this.Controls.Add(this.end);
            this.Controls.Add(this.Show);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.FILENAME);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lab);
            this.Controls.Add(this.labelBox);
            this.Controls.Add(this.SearchWord);
            this.Controls.Add(this.Start);
            this.Name = "Form1";
            this.Text = "Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.TextBox SearchWord;
        private System.Windows.Forms.ComboBox labelBox;
        private System.Windows.Forms.Label lab;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label FILENAME;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox Show;
        private System.Windows.Forms.Button end;
    }
}

