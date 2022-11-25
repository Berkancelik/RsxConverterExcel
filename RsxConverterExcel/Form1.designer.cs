namespace ResourceToExcelConveretr
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
            this.btnSourceBrowse = new System.Windows.Forms.Button();
            this.txtSourcePath = new System.Windows.Forms.TextBox();
            this.txtTargetPath = new System.Windows.Forms.TextBox();
            this.btnTargetBrowse = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.chkGenerateSingle = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnSourceBrowse
            // 
            this.btnSourceBrowse.Location = new System.Drawing.Point(336, 39);
            this.btnSourceBrowse.Name = "btnSourceBrowse";
            this.btnSourceBrowse.Size = new System.Drawing.Size(64, 23);
            this.btnSourceBrowse.TabIndex = 0;
            this.btnSourceBrowse.Text = "Browse";
            this.btnSourceBrowse.UseVisualStyleBackColor = true;
            this.btnSourceBrowse.Click += new System.EventHandler(this.btnSourceBrowse_Click);
            // 
            // txtSourcePath
            // 
            this.txtSourcePath.Location = new System.Drawing.Point(12, 40);
            this.txtSourcePath.Margin = new System.Windows.Forms.Padding(0);
            this.txtSourcePath.Name = "txtSourcePath";
            this.txtSourcePath.ReadOnly = true;
            this.txtSourcePath.Size = new System.Drawing.Size(311, 20);
            this.txtSourcePath.TabIndex = 1;
            // 
            // txtTargetPath
            // 
            this.txtTargetPath.Location = new System.Drawing.Point(12, 105);
            this.txtTargetPath.Name = "txtTargetPath";
            this.txtTargetPath.ReadOnly = true;
            this.txtTargetPath.Size = new System.Drawing.Size(311, 20);
            this.txtTargetPath.TabIndex = 2;
            // 
            // btnTargetBrowse
            // 
            this.btnTargetBrowse.Location = new System.Drawing.Point(336, 103);
            this.btnTargetBrowse.Name = "btnTargetBrowse";
            this.btnTargetBrowse.Size = new System.Drawing.Size(64, 23);
            this.btnTargetBrowse.TabIndex = 1;
            this.btnTargetBrowse.Text = "Browse";
            this.btnTargetBrowse.UseVisualStyleBackColor = true;
            this.btnTargetBrowse.Click += new System.EventHandler(this.btnTargetBrowse_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(168, 250);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(91, 38);
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Path of source folder:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Path of target folder:";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(15, 207);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(385, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 191);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(157, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Wait till the operation completes";
            // 
            // chkGenerateSingle
            // 
            this.chkGenerateSingle.AutoSize = true;
            this.chkGenerateSingle.Location = new System.Drawing.Point(15, 155);
            this.chkGenerateSingle.Name = "chkGenerateSingle";
            this.chkGenerateSingle.Size = new System.Drawing.Size(144, 17);
            this.chkGenerateSingle.TabIndex = 10;
            this.chkGenerateSingle.Text = "Generate single excel file";
            this.chkGenerateSingle.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(427, 346);
            this.Controls.Add(this.chkGenerateSingle);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnTargetBrowse);
            this.Controls.Add(this.txtTargetPath);
            this.Controls.Add(this.txtSourcePath);
            this.Controls.Add(this.btnSourceBrowse);
            this.Name = "Form1";
            this.Text = "Resources to Excel converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSourceBrowse;
        private System.Windows.Forms.TextBox txtSourcePath;
        private System.Windows.Forms.TextBox txtTargetPath;
        private System.Windows.Forms.Button btnTargetBrowse;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkGenerateSingle;
    }
}

