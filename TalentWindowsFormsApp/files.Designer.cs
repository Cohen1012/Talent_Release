namespace TalentWindowsFormsApp
{
    partial class Files
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.CloseBtn = new System.Windows.Forms.Button();
            this.SaveFilesBtn = new System.Windows.Forms.Button();
            this.NewFileBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.AutoScroll = true;
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.CloseBtn);
            this.splitContainer1.Panel2.Controls.Add(this.SaveFilesBtn);
            this.splitContainer1.Panel2.Controls.Add(this.NewFileBtn);
            this.splitContainer1.Size = new System.Drawing.Size(377, 269);
            this.splitContainer1.SplitterDistance = 215;
            this.splitContainer1.TabIndex = 0;
            // 
            // CloseBtn
            // 
            this.CloseBtn.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.CloseBtn.Location = new System.Drawing.Point(299, 13);
            this.CloseBtn.Name = "CloseBtn";
            this.CloseBtn.Size = new System.Drawing.Size(75, 25);
            this.CloseBtn.TabIndex = 2;
            this.CloseBtn.Text = "離開";
            this.CloseBtn.UseVisualStyleBackColor = true;
            this.CloseBtn.Click += new System.EventHandler(this.CloseBtn_Click);
            // 
            // SaveFilesBtn
            // 
            this.SaveFilesBtn.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.SaveFilesBtn.Location = new System.Drawing.Point(218, 13);
            this.SaveFilesBtn.Name = "SaveFilesBtn";
            this.SaveFilesBtn.Size = new System.Drawing.Size(75, 25);
            this.SaveFilesBtn.TabIndex = 1;
            this.SaveFilesBtn.Text = "送出";
            this.SaveFilesBtn.UseVisualStyleBackColor = true;
            this.SaveFilesBtn.Click += new System.EventHandler(this.SaveFilesBtn_Click);
            // 
            // NewFileBtn
            // 
            this.NewFileBtn.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.NewFileBtn.Location = new System.Drawing.Point(13, 13);
            this.NewFileBtn.Name = "NewFileBtn";
            this.NewFileBtn.Size = new System.Drawing.Size(75, 25);
            this.NewFileBtn.TabIndex = 0;
            this.NewFileBtn.Text = "新增";
            this.NewFileBtn.UseVisualStyleBackColor = true;
            this.NewFileBtn.Click += new System.EventHandler(this.NewFileBtn_Click);
            // 
            // Files
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 269);
            this.Controls.Add(this.splitContainer1);
            this.Name = "Files";
            this.Text = "files";
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button NewFileBtn;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.Button SaveFilesBtn;
    }
}