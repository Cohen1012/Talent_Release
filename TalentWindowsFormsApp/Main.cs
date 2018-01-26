using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary;

namespace TalentWindowsFormsApp
{
    public partial class Main : Form
    {
        /// <summary>
        /// 帳號
        /// </summary>
        private string Account = string.Empty;
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            this.ShowSignIn();
        }

        /// <summary>
        /// 執行登入功能
        /// </summary>
        private void ShowSignIn()
        {
            SignIn signIn = new SignIn();
            this.Hide();
            this.DialogResult = signIn.ShowDialog();
            if (this.DialogResult == DialogResult.OK)
            {
                Account = signIn.Msg;
                this.Text = signIn.Msg;
                this.Show();
                signIn.Close();
                if (HaveOpened(this, "AuthorityManagement"))
                {
                    AuthorityManagement authorityManagement = new AuthorityManagement(signIn.Msg)
                    {
                        MdiParent = this
                    };
                    authorityManagement.Show();
                }
            }
            else if (this.DialogResult == DialogResult.Cancel)
            {
                this.Close();
            }
        }

        /// <summary>
        /// 登出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SignOutBtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.MdiChildren.Length; i++)
            {
                this.MdiChildren[i].Close();
            }
            this.ShowSignIn();
        }

        /// <summary>
        /// 開權限管理功能視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AuthorityBtn_Click(object sender, EventArgs e)
        {
            if (HaveOpened(this, "AuthorityManagement"))
            {
                AuthorityManagement authorityManagement = new AuthorityManagement(Account)
                {
                    MdiParent = this
                };
                CloseFrom("AuthorityManagement");
                authorityManagement.Show();
            }
        }

        /// <summary>
        /// 限制MDI子表單不能重複開啟 
        /// </summary>
        /// <param name="main">父表單</param>
        /// <param name="childName">子表單Name</param>
        /// <returns></returns>
        private bool HaveOpened(Main main, string childName)
        {
            //查看視窗是否已經被打開
            bool bReturn = true;
            for (int i = 0; i < main.MdiChildren.Length; i++)
            {
                if (main.MdiChildren[i].Name == childName)
                {
                    main.MdiChildren[i].BringToFront();
                    bReturn = false;
                    break;
                }
            }
            return bReturn;
        }

        /// <summary>
        /// 開啟資料管理視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataBtn_Click(object sender, EventArgs e)
        {
            if (HaveOpened(this, "TalentSearch"))
            {
                TalentSearch talentSearch = new TalentSearch
                {
                    MdiParent = this
                };
                CloseFrom("TalentSearch");
                talentSearch.Show();
            }
        }

        /// <summary>
        /// 關閉正在執行以外的視窗
        /// </summary>
        /// <param name="formName"></param>
        private void CloseFrom(string formName)
        {
            for (int i = 0; i < this.MdiChildren.Length; i++)
            {
                if (this.MdiChildren[i].Name != formName)
                {
                    this.MdiChildren[i].Close();
                }
            }
        }
    }
}
