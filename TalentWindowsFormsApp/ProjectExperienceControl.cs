using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary.Model;
using TalentClassLibrary;

namespace TalentWindowsFormsApp
{
    public partial class ProjectExperienceControl : UserControl
    {
        /// <summary>
        /// 觸發事件
        /// </summary>
        public event EventHandler BtnClick;

        /// <summary>
        /// 觸發刪除事件
        /// </summary>
        /// <param name="e"></param>
        protected void OnbtnClick(EventArgs e)
        {
            BtnClick?.Invoke(this, e);
        }

        public ProjectExperienceControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 將專案經驗顯示出來
        /// </summary>
        /// <param name="project"></param>
        public void SetProjectExperience(ProjectExperience project)
        {
            CompanyTxt.Text = project.Company;
            ProjectNameTxt.Text = project.Project_Name;
            OSTxt.Text = project.OS;
            DatabaseTxt.Text = project.Database;
            PositionTxt.Text = project.Position;
            StartEndDateTxt.Text = project.Start_End_Date;
            LanguageTxt.Text = project.Language;
            ToolsTxt.Text = project.Tools;
            DescriptionTxt.Text = project.Description;
        }

        public ProjectExperience GetProjectExperience()
        {
            ProjectExperience project = new ProjectExperience
            {
                Company = CompanyTxt.Text,
                Project_Name = ProjectNameTxt.Text,
                OS = OSTxt.Text,
                Database = DatabaseTxt.Text,
                Position = PositionTxt.Text,
                Start_End_Date = StartEndDateTxt.Text,
                Language = LanguageTxt.Text,
                Tools = ToolsTxt.Text,
                Description = DescriptionTxt.Text,
            };

            return project;
        }

        /// <summary>
        /// 刪除事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                // 觸發 btnClick Event
                OnbtnClick(e);
            }
        }
    }
}
