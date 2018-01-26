using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TalentClassLibrary.Model;
using ServiceStack.Text;
using ServiceStack;

namespace TalentClassLibrary.Tests
{
    class TestData
    {
        private static TestData testData = new TestData();

        public static TestData GetInstance() => testData;

        public DataTable TestExportData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("姓名");
            dt.Columns.Add("職缺");
            dt.Columns.Add("面談日期");
            dt.Columns.Add("最後編輯時間");
            dt.Rows.Add("賴志維", "軟體工程師", "106/08/24", "106/08/21");
            dt.Rows.Add("李忠鍵", "軟體工程師", "106/12/29", "106/12/12");
            return dt;
        }

        public DataTable TestContactSituationData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Contact_Date");
            dt.Columns.Add("Contact_Status");
            dt.Columns.Add("Remarks");
            dt.Rows.Add("1", "2017/08/16", "104邀約", "C#");
            dt.Rows.Add("2", "2017/12/01", "暫不考慮", "JAVA");
            return dt;
        }

        public DataTable TestUpdateContactSituationData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Contact_status_Id");
            dt.Columns.Add("Contact_Date");
            dt.Columns.Add("Contact_Status");
            dt.Columns.Add("Remarks");
            dt.Rows.Add("1", DateTime.Now.ToString(), "104邀約", "C#");
            dt.Rows.Add("2", DateTime.Now.ToString(), "電話聯繫", "JAVA");
            return dt;
        }

        public DataTable TestDelContactSituationData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Contact_status_Id");
            dt.Columns.Add("Contact_Date");
            dt.Columns.Add("Contact_Status");
            dt.Columns.Add("Remarks");
            dt.Rows.Add("3", "2017/08/17", "104邀約", "C#");
            dt.Rows.Add("4", "2017/12/01", "電話聯繫", "JAVA");
            return dt;
        }

        public DataTable TestContactSituationActionData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Contact_status_Id");
            dt.Columns.Add("Contact_Date");
            dt.Columns.Add("Contact_Status");
            dt.Columns.Add("Remarks");
            dt.Columns.Add("Flag");
            dt.Rows.Add("1", DateTime.Now.ToString(), "104邀約", "C#", "U");
            dt.Rows.Add("9", DateTime.Now.ToString(), "電話聯繫", "JAVA", "D");
            dt.Rows.Add("9", DateTime.Now.ToString(), "人才儲存", "JAVA", "I");
            return dt;
        }

        public DataTable TestContactSituationInfoData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Sex");
            dt.Columns.Add("Mail");
            dt.Columns.Add("CellPhone");
            dt.Columns.Add("Place");
            dt.Columns.Add("Skill");
            dt.Columns.Add("Cooperation_Mode");
            dt.Columns.Add("Status");
            dt.Columns.Add("UpdateTime");
            dt.Rows.Add("賴志維", "男", "monkey60146@yahoo.com.tw", "0978535528", "雲林縣", "MS SQL、Asp.Net、C#", "全職", "追蹤", DateTime.Now.ToString());
            dt.Rows.Add("李忠鍵", "男", "monkey60146@yahoo.com.tw", "0965438536", "彰化縣", "MS SQL、Asp.Net、C#、JAVA", "合約", "保留", DateTime.Now.ToString());
            return dt;
        }

        public DataTable TestInterviewData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Code");
            dt.Columns.Add("Sex");
            dt.Columns.Add("Mail");
            dt.Columns.Add("CellPhone");
            dt.Columns.Add("Picture");
            dt.Columns.Add("Vacancies");
            dt.Columns.Add("Married");
            dt.Columns.Add("Interview_Date");
            dt.Columns.Add("Birthday");
            dt.Rows.Add("賴志維", "0203", "男", "monkey60146@yahoo.com.tw", "0978535528", "./images/edit.png", "軟體工程師", "未婚", "2017/08/24", "1995/06/16");
            dt.Rows.Add("李忠鍵", "0204", "男", "monkey60146@yahoo.com.tw", "0965438536", "./images/delete.png", "軟體工程師", "未婚", "2017/12/29", "1995/06/14");
            return dt;
        }

        public DataTable TestInsertInterviewInfoData()
        {
            DataTable dt = this.CreateInterviewInfoTable();
            dt.Rows.Add(this.CreateEducationData(), this.CreateLanguageData(), this.CreateWorkExperienceData(), "軟體工程師", "2017/12/29", "賴志維");
            return dt;
        }

        private string CreateEducationData()
        {
            List<Education> educationList = new List<Education>();
            Education education = new Education
            {
                School = "斗六家商",
                Department = "資料處理科"
            };
            educationList.Add(education);
            string educationJson = educationList.ToJson();
            return educationJson;
        }

        private string CreateLanguageData()
        {
            List<Language> languageList = new List<Language>();
            Language language = new Language
            {
                Language_Name = "英文",
                Listen = "平",
                Read = "平",
                Speak = "平",
                Write = "平"
            };
            languageList.Add(language);
            string languageJson = languageList.ToJson();
            return languageJson;
        }

        private string CreateWorkExperienceData()
        {
            List<WorkExperience> workExperienceList = new List<WorkExperience>();
            WorkExperience language = new WorkExperience
            {
                Institution_name = "屏東科技大學",
                Position = "學生",
                Leaving_Reason = "畢業"
            };
            workExperienceList.Add(language);
            string workExperienceJson = workExperienceList.ToJson();
            return workExperienceJson;
        }

        public DataTable TestUpdateInterviewInfoData()
        {
            DataTable dt = this.CreateInterviewInfoTable();
            dt.Rows.Add(this.CreateEducationData(), this.CreateLanguageData(), this.CreateWorkExperienceData(), "軟體工程師", "2017/12/29", "賴志維", "男");
            return dt;
        }

        public DataTable TestProjectExperienceData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Interview_Id");
            dt.Columns.Add("Company");
            dt.Columns.Add("Project_Name");
            dt.Columns.Add("OS");
            dt.Columns.Add("Database");
            dt.Columns.Add("Position");
            dt.Columns.Add("Start_End_Date");
            dt.Columns.Add("Language");
            dt.Columns.Add("Tools");
            dt.Columns.Add("Description");
            dt.Rows.Add("2", "屏東科技大學", "智慧水產養殖管理系統", "Windows 10", "MS SQL", "學生", "105/07/01~105/12/14", "C#", "Asp.Net", "say somthing");
            dt.Rows.Add("2", "亦思科技股份有限公司", "人才資料庫系統", "Windows 10", "MS SQL", "軟體工程師", "106/12/11~106/12/29", "C#", "Asp.Net", "say somthing");
            return dt;
        }

        public DataTable TestProjectExperienceData1()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Interview_Id");
            dt.Columns.Add("Company");
            dt.Columns.Add("Project_Name");
            dt.Columns.Add("OS");
            dt.Columns.Add("Database");
            dt.Columns.Add("Position");
            dt.Columns.Add("Start_End_Date");
            dt.Columns.Add("Language");
            dt.Columns.Add("Tools");
            dt.Columns.Add("Description");
            dt.Rows.Add("2", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            dt.Rows.Add("2", "亦思科技股份有限公司", "人才資料庫系統", "Windows 10", "MS SQL", "軟體工程師", "106/12/11~106/12/29", "C#", "Asp.Net", "say somthing");
            return dt;
        }

        public DataSet TestSaveInterviewResultData()
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Appointment");
            dt.Columns.Add("Results_Remark");
            dt.Rows.Add("錄用", string.Empty);
            ds.Tables.Add(dt);
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Interviewer");
            dt1.Columns.Add("Result");
            dt1.Rows.Add("國基", "還有很多東西要學習");
            ds.Tables.Add(dt1);
            return ds;
        }

        public DataTable TestProjectExperienceData(string[] testData1, string[] testData2)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Interview_Id");
            dt.Columns.Add("Company");
            dt.Columns.Add("Project_Name");
            dt.Columns.Add("OS");
            dt.Columns.Add("Database");
            dt.Columns.Add("Position");
            dt.Columns.Add("Starting_Date");
            dt.Columns.Add("Closing_Date");
            dt.Columns.Add("Language");
            dt.Columns.Add("Tools");
            dt.Columns.Add("Description");
            dt.Rows.Add(testData1);
            dt.Rows.Add(testData2);
            return dt;
        }

        public DataTable TestCodeData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Flag");
            dt.Rows.Add("60146", "1", "I");
            dt.Rows.Add("913736", "1", "I");
            return dt;
        }

        public DataTable TestCodeData2()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Flag");
            dt.Rows.Add("60146", "1", "U");
            dt.Rows.Add("913736", "1", "U");
            dt.Rows.Add("TEST", "1", "I");
            return dt;
        }

        /// <summary>
        /// 給Action用的測試資料
        /// </summary>
        /// <returns></returns>
        public DataTable TestCodeData3()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Flag");
            dt.Rows.Add("60146", "1", "D");
            dt.Rows.Add("913736", "1", "U");
            dt.Rows.Add("TEST", "1", "I");
            dt.Rows.Add("TEST1", "1", "");
            return dt;
        }

        public DataTable TestCodeActionData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Flag");
            dt.Rows.Add("", "0205", "1", "D");
            dt.Rows.Add("2", "60146", "1", "U");
            dt.Rows.Add("", "TEST", "2", "I");
            dt.Rows.Add("", "TEST1", "2", "");
            return dt;
        }

        public DataTable TestCodeActionData1()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Flag");
            dt.Rows.Add("", "0205", "1", "D");
            dt.Rows.Add("2", "60146", "1", "U");
            dt.Rows.Add("", "60146", "2", "I");
            dt.Rows.Add("", "TEST1", "2", "");
            return dt;
        }

        public DataTable TestEducationData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("School");
            dt.Columns.Add("Department");
            dt.Columns.Add("Start_End_Date");
            dt.Columns.Add("Is_Graduation");
            dt.Columns.Add("Remark");
            dt.Rows.Add("屏東科技大學", "資訊管理系", "", "", "");
            return dt;
        }

        public DataTable TestEducationData1()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("School1");
            dt.Columns.Add("Department");
            dt.Columns.Add("Start_End_Date");
            dt.Columns.Add("Is_Graduation");
            dt.Columns.Add("Remark");
            dt.Rows.Add("屏東科技大學", "資訊管理系", "", "", "");
            return dt;
        }

        public DataTable TestWorkExperienceData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Institution_name");
            dt.Columns.Add("Position");
            dt.Columns.Add("Start_End_Date");
            dt.Columns.Add("Start_Salary");
            dt.Columns.Add("End_Salary");
            dt.Columns.Add("Leaving_Reason");
            dt.Rows.Add("屏東科技大學", "學生", string.Empty, string.Empty, string.Empty, "畢業");
            return dt;
        }

        public DataTable TestWorkExperienceData1()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Institution_name1");
            dt.Columns.Add("Position");
            dt.Columns.Add("Start_End_Date");
            dt.Columns.Add("Start_Salary");
            dt.Columns.Add("End_Salary");
            dt.Columns.Add("Leaving_Reason");
            dt.Rows.Add("屏東科技大學", "學生", string.Empty, string.Empty, string.Empty, "畢業");
            return dt;
        }

        /// <summary>
        /// 創建InterviewInfo欄位
        /// </summary>
        /// <returns></returns>
        private DataTable CreateInterviewInfoTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Education");
            dt.Columns.Add("Language");
            dt.Columns.Add("Work_Experience");
            dt.Columns.Add("Vacancies");
            dt.Columns.Add("Interview_Date");
            dt.Columns.Add("Name");
            dt.Columns.Add("Sex");
            dt.Columns.Add("Birthday");
            dt.Columns.Add("Married");
            dt.Columns.Add("Mail");
            dt.Columns.Add("Adress");
            dt.Columns.Add("CellPhone");
            dt.Columns.Add("Image");
            dt.Columns.Add("Expertise_Language");
            dt.Columns.Add("Expertise_Tools");
            dt.Columns.Add("Expertise_Devops");
            dt.Columns.Add("Expertise_OS");
            dt.Columns.Add("Expertise_BigData");
            dt.Columns.Add("Expertise_DataBase");
            dt.Columns.Add("Expertise_Certification");
            dt.Columns.Add("IsStudy");
            dt.Columns.Add("IsService");
            dt.Columns.Add("Relatives_Relationship");
            dt.Columns.Add("Relatives_Name");
            dt.Columns.Add("Care_Work");
            dt.Columns.Add("Hope_Salary");
            dt.Columns.Add("When_Report");
            dt.Columns.Add("Advantage");
            dt.Columns.Add("Disadvantages");
            dt.Columns.Add("Hobby");
            dt.Columns.Add("Attract_Reason");
            dt.Columns.Add("Future_Goal");
            dt.Columns.Add("Hope_Supervisor");
            dt.Columns.Add("Hope_Promise");
            dt.Columns.Add("Introduction");
            dt.Columns.Add("During_Service");
            dt.Columns.Add("Exemption_Reason");
            dt.Columns.Add("Urgent_Contact_Person");
            dt.Columns.Add("Urgent_CellPhone");
            dt.Columns.Add("Urgent_Relationship");
            return dt;
        }

        public InterviewInfo TestExcelInterviewInfoData()
        {
            InterviewInfo interviewInfo = new InterviewInfo
            {
                Image = @".\images\member.jpg",
                Vacancies = "軟體工程師",
                Name = "賴志維",
                Married = "未婚",
                Adress = "雲林縣斗六市八德里文化路690-29號",
                Urgent_Contact_Person = "賴雲慈",
                Interview_Date = "2017/07/20",
                Sex = "男",
                Mail = "monkey60146@gmail.com",
                Urgent_Relationship = "父子",
                Birthday = "1995/06/16",
                CellPhone = "0978535528",
                Urgent_CellPhone = "0965438536"
            };

            ////學歷
            List<Education> educationList = new List<Education>();
            Education education = new Education()
            {
                School = "國立屏東科技大學",
                Department = "資訊管理系",
                Start_End_Date = "102/09~106/06",
                Is_Graduation = "畢業",
                Remark = "校內專題競賽第二名\n全國專題競賽佳作"
            };
            educationList.Add(education);
            Education education1 = new Education()
            {
                School = "國立斗六高級家事商業職業學校",
                Department = "資料處理科",
                Start_End_Date = "99/09~102/06",
                Is_Graduation = "畢業",
                Remark = string.Empty
            };
            educationList.Add(education1);
            interviewInfo.Education = educationList.ToJson();
            ////工作經驗
            List<WorkExperience> workExperienceList = new List<WorkExperience>();
            WorkExperience workExperience = new WorkExperience
            {
                Institution_name = "屏東科技大學",
                Position = "學生",
                Start_End_Date = "102/09~106/06",
                Start_Salary = "0",
                End_Salary = "0",
                Leaving_Reason = "畢業"
            };
            workExperienceList.Add(workExperience);
            interviewInfo.Work_Experience = workExperienceList.ToJson();
            //interviewInfo.Work_Experience = "[]";
            ////兵役
            interviewInfo.During_Service = string.Empty;
            interviewInfo.Exemption_Reason = "體重過輕";
            ////語言能力
            List<Language> languageList = new List<Language>();
            Language language = new Language
            {
                Language_Name = "英文",
                Listen = "平",
                Speak = "略",
                Read  = "平",
                Write = "略"
            };
            languageList.Add(language);
            interviewInfo.Language = languageList.ToJson();
            ////專長
            interviewInfo.Expertise_Language = "C#,R,Android";
            interviewInfo.Expertise_Tools = "Visual Studio,Eclipse,R Studio,Framwork：ASP.NET";
            interviewInfo.Expertise_Devops = "Git";
            interviewInfo.Expertise_OS = "Windows,Android";
            interviewInfo.Expertise_BigData = string.Empty;
            interviewInfo.Expertise_DataBase = "Oracle,MS SQL,QQ";
            interviewInfo.Expertise_Certification = string.Empty;


            ////問題
            interviewInfo.IsStudy = "無";
            interviewInfo.IsService = "無";
            interviewInfo.Relatives_Relationship = "無";
            interviewInfo.Relatives_Name = "無";
            interviewInfo.Care_Work = "無";
            interviewInfo.Hope_Salary = "無";
            interviewInfo.When_Report = "無";
            interviewInfo.Advantage = "無";
            interviewInfo.Disadvantages = "無";
            interviewInfo.Hobby = "無";
            interviewInfo.Attract_Reason = "無";
            interviewInfo.Future_Goal = "無";
            interviewInfo.Hope_Supervisor = "無";
            interviewInfo.Hope_Promise = "無";
            interviewInfo.Introduction = "無";

            return interviewInfo;
        }

        public InterviewResults TestExcelInterviewResultData()
        {
            InterviewResults interviewResults = new InterviewResults();
            InterviewResult interviewResult = new InterviewResult
            {
                Appointment = "錄用",
                Results_Remark = "say\nsomthing\nsay\nsomthing\nsay\nsomthing\nsay\nsomthing"
            };
            interviewResults.InterviewResult = interviewResult;

            List<InterviewComments> interviewCommentsList = new List<InterviewComments>();
            InterviewComments interviewComments = new InterviewComments
            {
                Interviewer = "國基",
                Result = "嘎U"
            };
            interviewCommentsList.Add(interviewComments);
            interviewCommentsList.Add(interviewComments);

            interviewResults.InterviewCommentsList = interviewCommentsList;

            return interviewResults;
        }

        public List<ProjectExperience> TestExcelProjectExperienceData()
        {
            List<ProjectExperience> projectExperienceList = new List<ProjectExperience>();
            ProjectExperience projectExperience = new ProjectExperience
            {
                Company = "屏東科技大學",
                Position = "學生",
                Project_Name = "智慧水產養殖管理資訊系統",
                Start_End_Date = "105/6/30~105/12/14",
                OS = "win 10",
                Language = "C#",
                Database = "MS SQL",
                Tools = "Visual Studio",
                Description = "say something"
            };
            projectExperienceList.Add(projectExperience);

            ProjectExperience projectExperience1 = new ProjectExperience
            {
                Company = string.Empty,
                Position = string.Empty,
                Project_Name = string.Empty,
                Start_End_Date = string.Empty,
                OS = string.Empty,
                Language = string.Empty,
                Database = string.Empty,
                Tools = string.Empty,
                Description = string.Empty
            };
            projectExperienceList.Add(projectExperience1);

            return projectExperienceList;
        }

        public List<ContactSituation> TestExcelContactSituationData()
        {
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            ContactInfo contactInfo = new ContactInfo
            {
                Name = "賴志維",
                Sex = "男",
                Mail = "monkey60146@gmail.com",
                CellPhone = "0978535528",
                Cooperation_Mode = "不限",
                Status = "追蹤",
                Place = "中科",
                Skill = "C#"
            };

            List<ContactStatus> contactStatusList = new List<ContactStatus>();
            for (int i = 0; i < 5; i++)
            {
                Models models = new Models();
                ContactStatus contactStatus = new ContactStatus
                {
                    Contact_Date = DateTime.Now.ToString("yyyy/MM/dd"),
                    Contact_Status = models.Contact_Status[i],
                    Remarks = "remarks"
                };

                contactStatusList.Add(contactStatus);
            }

            ContactSituation contactSituation = new ContactSituation
            {
                Info = contactInfo,
                Status = contactStatusList,
                Code = "60146\n913736\n10256048"
            };

            contactSituationList.Add(contactSituation);
            contactSituationList.Add(contactSituation);
            return contactSituationList;
        }
    }
}
