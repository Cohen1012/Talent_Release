namespace TalentWindowsFormsApp
{
    partial class FileControl
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.OpenFileDialogBtn = new System.Windows.Forms.Button();
            this.FileNameTxt = new System.Windows.Forms.TextBox();
            this.DowndloadFileLink = new System.Windows.Forms.LinkLabel();
            this.DelFileBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OpenFileDialogBtn
            // 
            this.OpenFileDialogBtn.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.OpenFileDialogBtn.Location = new System.Drawing.Point(4, 4);
            this.OpenFileDialogBtn.Name = "OpenFileDialogBtn";
            this.OpenFileDialogBtn.Size = new System.Drawing.Size(114, 26);
            this.OpenFileDialogBtn.TabIndex = 0;
            this.OpenFileDialogBtn.Text = "請選擇檔案";
            this.OpenFileDialogBtn.UseVisualStyleBackColor = true;
            this.OpenFileDialogBtn.Click += new System.EventHandler(this.OpenFileDialogBtn_Click);
            // 
            // FileNameTxt
            // 
            this.FileNameTxt.Enabled = false;
            this.FileNameTxt.Location = new System.Drawing.Point(123, 8);
            this.FileNameTxt.Margin = new System.Windows.Forms.Padding(2);
            this.FileNameTxt.Name = "FileNameTxt";
            this.FileNameTxt.Size = new System.Drawing.Size(114, 22);
            this.FileNameTxt.TabIndex = 1;
            // 
            // DowndloadFileLink
            // 
            this.DowndloadFileLink.AutoSize = true;
            this.DowndloadFileLink.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.DowndloadFileLink.Location = new System.Drawing.Point(243, 8);
            this.DowndloadFileLink.Name = "DowndloadFileLink";
            this.DowndloadFileLink.Size = new System.Drawing.Size(47, 19);
            this.DowndloadFileLink.TabIndex = 2;
            this.DowndloadFileLink.TabStop = true;
            this.DowndloadFileLink.Text = "下載";
            this.DowndloadFileLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.DowndloadFileLink_LinkClicked);
            // 
            // DelFileBtn
            // 
            this.DelFileBtn.Image = global::TalentWindowsFormsApp.Properties.Resources.cancel;
            this.DelFileBtn.Location = new System.Drawing.Point(297, 6);
            this.DelFileBtn.Name = "DelFileBtn";
            this.DelFileBtn.Size = new System.Drawing.Size(30, 30);
            this.DelFileBtn.TabIndex = 3;
            this.DelFileBtn.UseVisualStyleBackColor = true;
            this.DelFileBtn.Click += new System.EventHandler(this.DelFileBtn_Click);
            // 
            // FileControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.DelFileBtn);
            this.Controls.Add(this.DowndloadFileLink);
            this.Controls.Add(this.FileNameTxt);
            this.Controls.Add(this.OpenFileDialogBtn);
            this.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Name = "FileControl";
            this.Size = new System.Drawing.Size(330, 39);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OpenFileDialogBtn;
        private System.Windows.Forms.TextBox FileNameTxt;
        private System.Windows.Forms.LinkLabel DowndloadFileLink;
        private System.Windows.Forms.Button DelFileBtn;
    }
}
