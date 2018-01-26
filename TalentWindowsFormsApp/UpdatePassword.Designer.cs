namespace TalentWindowsFormsApp
{
    partial class UpdatePassword
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
            this.label1 = new System.Windows.Forms.Label();
            this.OldPasswordTxt = new System.Windows.Forms.TextBox();
            this.NewPasswordTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.CheckNewPasswordTxt = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.OKBtn = new System.Windows.Forms.Button();
            this.CancelBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(13, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "舊密碼：";
            // 
            // OldPasswordTxt
            // 
            this.OldPasswordTxt.Location = new System.Drawing.Point(180, 57);
            this.OldPasswordTxt.Name = "OldPasswordTxt";
            this.OldPasswordTxt.PasswordChar = '*';
            this.OldPasswordTxt.Size = new System.Drawing.Size(100, 22);
            this.OldPasswordTxt.TabIndex = 1;
            this.OldPasswordTxt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.OldPasswordTxt_KeyPress);
            // 
            // NewPasswordTxt
            // 
            this.NewPasswordTxt.Location = new System.Drawing.Point(180, 99);
            this.NewPasswordTxt.Name = "NewPasswordTxt";
            this.NewPasswordTxt.PasswordChar = '*';
            this.NewPasswordTxt.Size = new System.Drawing.Size(100, 22);
            this.NewPasswordTxt.TabIndex = 3;
            this.NewPasswordTxt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NewPasswordTxt_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(13, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 19);
            this.label2.TabIndex = 2;
            this.label2.Text = "新密碼：";
            // 
            // CheckNewPasswordTxt
            // 
            this.CheckNewPasswordTxt.Location = new System.Drawing.Point(180, 141);
            this.CheckNewPasswordTxt.Name = "CheckNewPasswordTxt";
            this.CheckNewPasswordTxt.PasswordChar = '*';
            this.CheckNewPasswordTxt.Size = new System.Drawing.Size(100, 22);
            this.CheckNewPasswordTxt.TabIndex = 5;
            this.CheckNewPasswordTxt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CheckNewPasswordTxt_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(13, 144);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(161, 19);
            this.label3.TabIndex = 4;
            this.label3.Text = "再次輸入新密碼：";
            // 
            // OKBtn
            // 
            this.OKBtn.Location = new System.Drawing.Point(99, 182);
            this.OKBtn.Name = "OKBtn";
            this.OKBtn.Size = new System.Drawing.Size(75, 23);
            this.OKBtn.TabIndex = 6;
            this.OKBtn.Text = "確定";
            this.OKBtn.UseVisualStyleBackColor = true;
            this.OKBtn.Click += new System.EventHandler(this.OKBtn_Click);
            // 
            // CancelBtn
            // 
            this.CancelBtn.Location = new System.Drawing.Point(204, 182);
            this.CancelBtn.Name = "CancelBtn";
            this.CancelBtn.Size = new System.Drawing.Size(75, 23);
            this.CancelBtn.TabIndex = 7;
            this.CancelBtn.Text = "取消";
            this.CancelBtn.UseVisualStyleBackColor = true;
            this.CancelBtn.Click += new System.EventHandler(this.CancelBtn_Click);
            // 
            // UpdatePassword
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(354, 261);
            this.Controls.Add(this.CancelBtn);
            this.Controls.Add(this.OKBtn);
            this.Controls.Add(this.CheckNewPasswordTxt);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.NewPasswordTxt);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.OldPasswordTxt);
            this.Controls.Add(this.label1);
            this.Name = "UpdatePassword";
            this.Text = "修改密碼";
            this.Load += new System.EventHandler(this.UpdatePassword_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox OldPasswordTxt;
        private System.Windows.Forms.TextBox NewPasswordTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox CheckNewPasswordTxt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button OKBtn;
        private System.Windows.Forms.Button CancelBtn;
    }
}