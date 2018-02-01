using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalentClassLibrary.Model
{
    public class Models
    {
        public List<string> Contact_Status = new List<string>() { "(無)", "人才儲存", "1111邀約", "104邀約", "mail邀約", "感謝函", "電話聯繫", "電聯未接", "技術訪談", "暫不考慮", "信件聯繫", "婉拒邀約", "同意邀約", "詢問問題", "主動應徵", "取消面談", "面談未到", "關閉履歷", "履歷部分公開", "人資系統資料" };
    }

    /// <summary>
    /// 指定是哪個資料的附加檔案
    /// </summary>
    public enum AttachedFilesMode
    {
        /// <summary>
        /// 面談基本資料
        /// </summary>
        InterviewInfo = 0,

        /// <summary>
        /// 專案經驗
        /// </summary>
        ProjectExperience = 1,
    }

    /// <summary>
    /// 學歷
    /// </summary>
    public class Education
    {
        public string School { get; set; }
        public string Department { get; set; }
        public string Start_End_Date { get; set; }
        public string Is_Graduation { get; set; }
        public string Remark { get; set; }
    }

    /// <summary>
    /// 語言能力
    /// </summary>
    public class Language
    {
        public string Language_Name { get; set; }
        public string Listen { get; set; }
        public string Speak { get; set; }
        public string Read { get; set; }
        public string Write { get; set; }
    }

    /// <summary>
    /// 工作經驗
    /// </summary>
    public class WorkExperience
    {
        public string Institution_name { get; set; }
        public string Position { get; set; }
        public string Start_End_Date { get; set; }
        public string Start_Salary { get; set; }
        public string End_Salary { get; set; }
        public string Leaving_Reason { get; set; }
    }

    /// <summary>
    /// 查詢頁面Grid所要呈現的格式
    /// </summary>
    public class SearchResult
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public List<string> Code { get; set; }
        public List<ContactStatus> Status { get; set; }
        public List<string> Interview_Date { get; set; }
        public string UpdateTime { get; set; }
    }

    /// <summary>
    /// 聯繫狀況資料
    /// </summary>
    public class ContactSituation
    {
        public ContactInfo Info { get; set; }
        public List<ContactStatus> Status { get; set; }
        public string Code { get; set; }
    }

    /// <summary>
    /// 聯繫基本資料
    /// </summary>
    public class ContactInfo
    {
        /// <summary>
        /// 年次
        /// </summary>
        public string Year { get; set; }
        public string Name { get; set; }
        public string Sex { get; set; }
        public string Mail { get; set; }
        public string CellPhone { get; set; }
        public string Cooperation_Mode { get; set; }
        public string Status { get; set; }
        public string Place { get; set; }
        public string Skill { get; set; }
    }

    /// <summary>
    /// 聯繫狀況
    /// </summary>
    public class ContactStatus
    {
        public string Contact_Date { get; set; }
        public string Contact_Status { get; set; }
        public string Remarks { get; set; }
    }

    /// <summary>
    /// 面談資料
    /// </summary>
    public class InterviewData
    {
        public InterviewInfo InterviewInfo { get; set; }
        public List<ProjectExperience> ProjectExperienceList { get; set; }
        public InterviewResults InterviewResults { get; set; }
    }

    /// <summary>
    /// 面談基本資料
    /// </summary>
    public class InterviewInfo
    {
        public string Vacancies { get; set; }
        public string Interview_Date { get; set; }
        public string Name { get; set; }
        public string Sex { get; set; }
        public string Birthday { get; set; }
        public string Married { get; set; }
        public string Mail { get; set; }
        public string Adress { get; set; }
        public string CellPhone { get; set; }
        public string Image { get; set; }
        public string Expertise_Language { get; set; }
        public string Expertise_Tools { get; set; }
        public string Expertise_Tools_Framwork { get; set; }
        public string Expertise_Devops { get; set; }
        public string Expertise_OS { get; set; }
        public string Expertise_BigData { get; set; }
        public string Expertise_DataBase { get; set; }
        public string Expertise_Certification { get; set; }
        public string IsStudy { get; set; }
        public string IsService { get; set; }
        public string Relatives_Relationship { get; set; }
        public string Relatives_Name { get; set; }
        public string Care_Work { get; set; }
        public string Hope_Salary { get; set; }
        public string When_Report { get; set; }
        public string Advantage { get; set; }
        public string Disadvantages { get; set; }
        public string Hobby { get; set; }
        public string Attract_Reason { get; set; }
        public string Future_Goal { get; set; }
        public string Hope_Supervisor { get; set; }
        public string Hope_Promise { get; set; }
        public string Introduction { get; set; }
        public string During_Service { get; set; }
        public string Exemption_Reason { get; set; }
        public string Urgent_Contact_Person { get; set; }
        public string Urgent_Relationship { get; set; }
        public string Urgent_CellPhone { get; set; }
        public string Education { get; set; }
        public string Language { get; set; }
        public string Work_Experience { get; set; }
    }

    /// <summary>
    /// 面談結果(任用評定與備註)
    /// </summary>
    public class InterviewResult
    {
        public string Appointment { get; set; }
        public string Results_Remark { get; set; }
    }

    /// <summary>
    /// 面談評語
    /// </summary>
    public class InterviewComments
    {
        public string Interviewer { get; set; }
        public string Result { get; set; }
    }

    /// <summary>
    /// 面談結果(Excel)
    /// </summary>
    public class InterviewResults
    {
        public InterviewResult InterviewResult { get; set; }
        public  List<InterviewComments> InterviewCommentsList { get; set; }
    }

    /// <summary>
    /// 專案經驗
    /// </summary>
    public class ProjectExperience
    {
        public string Company { get; set; }
        public string Project_Name { get; set; }
        public string OS { get; set; }
        public string Database { get; set; }
        public string Position { get; set; }
        public string Language { get; set; }
        public string Tools { get; set; }
        public string Description { get; set; }
        public string Start_End_Date { get; set; }
    }

    /// <summary>
    /// 附加檔案
    /// </summary>
    public class AttachedFiles
    {
        public string File_Path { get; set; }
        public string Belong { get; set; }
    }
}
