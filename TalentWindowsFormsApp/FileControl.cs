using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary;
using TalentClassLibrary.Model;
using System.IO;

namespace TalentWindowsFormsApp
{
    public partial class FileControl : UserControl
    {
        /// <summary>
        /// 路徑
        /// </summary>
        private string FilePath = string.Empty;

        /// <summary>
        /// 觸發事件
        /// </summary>
        public event EventHandler BtnClick;

        public event EventHandler FileFilterClick;

        /// <summary>
        /// 觸發刪除事件
        /// </summary>
        /// <param name="e"></param>
        protected void OnbtnClick(EventArgs e)
        {
            BtnClick?.Invoke(this, e);
        }

        /// <summary>
        /// 觸發選取檔案件
        /// </summary>
        /// <param name="e"></param>
        protected void OnFileFilterClick(EventArgs e)
        {
            FileFilterClick?.Invoke(this, e);
        }

        public FileControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 取得檔名
        /// </summary>
        /// <returns></returns>
        public string GetFileName()
        {
            return FileNameTxt.Text;
        }

        /// <summary>
        /// 取得路徑
        /// </summary>
        /// <returns></returns>
        public string GetFilePath()
        {
            return FilePath;
        }

        /// <summary>
        /// 儲存檔案資訊
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public void SetFileInfo(string path, string fileName)
        {
            FileNameTxt.Text = fileName;
            FilePath = path;
        }

        /// <summary>
        /// 取得檔案資訊
        /// </summary>
        /// <returns></returns>
        public AttachedFiles GetFileInfo()
        {
            AttachedFiles attachedFiles = new AttachedFiles
            {
                File_Path = FilePath
            };

            return attachedFiles;
        }

        /// <summary>
        /// 選取檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFileDialogBtn_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(FileNameTxt.Text))
            {
                DialogResult result = MessageBox.Show("警告舊檔案覆蓋後會消失，是否繼續", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    return;
                }
            }

            OnFileFilterClick(e);
        }

        private void DowndloadFileLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if(string.IsNullOrEmpty(FilePath))
            {
                MessageBox.Show("沒有檔案", "錯誤訊息");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "請選擇存檔的路徑",
                FileName = FileNameTxt.Text,
                Filter = @"All Image And PDF Files|*.bmp;*.ico;*.gif;*.jpeg;*.jpg;*.png;*.tif;*.tiff;*.pdf|BMP|*.bmp|ICO|*.ico|GIF|*.gif|JEPG|*.jpeg|JPG|*.jpg|PNG|*.png|TIF|*.tif|TIFF|*.tiff|PDF|*.pdf"
            };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string msg = Talent.GetInstance().DownloadFile(FilePath, sfd.FileName);
                MessageBox.Show(msg,"訊息");
            }
        }

        /// <summary>
        /// 刪除事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelFileBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否刪除選取的檔案", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                // 觸發 btnClick Event
                OnbtnClick(e);
            }
        }
    }
}
