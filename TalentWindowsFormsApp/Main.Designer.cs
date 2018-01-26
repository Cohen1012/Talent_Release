namespace TalentWindowsFormsApp
{
    partial class Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.AuthorityBtn = new System.Windows.Forms.ToolStripButton();
            this.DataBtn = new System.Windows.Forms.ToolStripButton();
            this.SignOutBtn = new System.Windows.Forms.ToolStripButton();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AuthorityBtn,
            this.DataBtn,
            this.SignOutBtn});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(284, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // AuthorityBtn
            // 
            this.AuthorityBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.AuthorityBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AuthorityBtn.Name = "AuthorityBtn";
            this.AuthorityBtn.Size = new System.Drawing.Size(59, 22);
            this.AuthorityBtn.Text = "權限管理";
            this.AuthorityBtn.Click += new System.EventHandler(this.AuthorityBtn_Click);
            // 
            // DataBtn
            // 
            this.DataBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.DataBtn.Image = ((System.Drawing.Image)(resources.GetObject("DataBtn.Image")));
            this.DataBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.DataBtn.Name = "DataBtn";
            this.DataBtn.Size = new System.Drawing.Size(59, 22);
            this.DataBtn.Text = "資料管理";
            this.DataBtn.Click += new System.EventHandler(this.DataBtn_Click);
            // 
            // SignOutBtn
            // 
            this.SignOutBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.SignOutBtn.Image = ((System.Drawing.Image)(resources.GetObject("SignOutBtn.Image")));
            this.SignOutBtn.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.SignOutBtn.Name = "SignOutBtn";
            this.SignOutBtn.Size = new System.Drawing.Size(35, 22);
            this.SignOutBtn.Text = "登出";
            this.SignOutBtn.Click += new System.EventHandler(this.SignOutBtn_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.toolStrip1);
            this.IsMdiContainer = true;
            this.Name = "Main";
            this.Text = "Main";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Main_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton AuthorityBtn;
        private System.Windows.Forms.ToolStripButton DataBtn;
        private System.Windows.Forms.ToolStripButton SignOutBtn;
    }
}