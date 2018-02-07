using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalentClassLibrary.Model
{
    public class Models
    {
        /// <summary>
        /// 聯繫裝況
        /// </summary>
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
        /// <summary>
        /// 學校名稱
        /// </summary>
        public string School { get; set; }
        /// <summary>
        /// 科系
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// 起迄年月
        /// </summary>
        public string Start_End_Date { get; set; }
        /// <summary>
        /// 畢/肄業
        /// </summary>
        public string Is_Graduation { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
    }

    /// <summary>
    /// 語言能力
    /// </summary>
    public class Language
    {
        /// <summary>
        /// 語言
        /// </summary>
        public string Language_Name { get; set; }
        /// <summary>
        /// 聽
        /// </summary>
        public string Listen { get; set; }
        /// <summary>
        /// 說
        /// </summary>
        public string Speak { get; set; }
        /// <summary>
        /// 讀
        /// </summary>
        public string Read { get; set; }
        /// <summary>
        /// 寫
        /// </summary>
        public string Write { get; set; }
    }

    /// <summary>
    /// 工作經驗
    /// </summary>
    public class WorkExperience
    {
        /// <summary>
        /// 機構名稱
        /// </summary>
        public string Institution_name { get; set; }
        /// <summary>
        /// 職稱
        /// </summary>
        public string Position { get; set; }
        /// <summary>
        /// 起迄年月
        /// </summary>
        public string Start_End_Date { get; set; }
        /// <summary>
        /// 到職薪水
        /// </summary>
        public string Start_Salary { get; set; }
        /// <summary>
        /// 離職薪水
        /// </summary>
        public string End_Salary { get; set; }
        /// <summary>
        /// 離職原因
        /// </summary>
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
        /// <summary>
        /// 聯繫基本資料
        /// </summary>
        public ContactInfo Info { get; set; }
        /// <summary>
        /// 聯繫狀況資料
        /// </summary>
        public List<ContactStatus> Status { get; set; }
        /// <summary>
        /// 代碼
        /// </summary>
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
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 性別
        /// </summary>
        public string Sex { get; set; }
        /// <summary>
        /// e-mail
        /// </summary>
        public string Mail { get; set; }
        /// <summary>
        /// 手機
        /// </summary>
        public string CellPhone { get; set; }
        /// <summary>
        /// 合作模式值為"全職"or"合約"or"皆可
        /// </summary>
        public string Cooperation_Mode { get; set; }
        /// <summary>
        /// 狀態值為"追蹤"or"保留"
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// 地點
        /// </summary>
        public string Place { get; set; }
        /// <summary>
        /// 技能
        /// </summary>
        public string Skill { get; set; }
    }

    /// <summary>
    /// 聯繫狀況
    /// </summary>
    public class ContactStatus
    {
        /// <summary>
        /// 聯繫日期
        /// </summary>
        public string Contact_Date { get; set; }
        /// <summary>
        /// 聯繫狀況
        /// </summary>
        public string Contact_Status { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remarks { get; set; }
    }

    /// <summary>
    /// 面談資料
    /// </summary>
    public class InterviewData
    {
        /// <summary>
        /// 面談基本資料
        /// </summary>
        public InterviewInfo InterviewInfo { get; set; }
        /// <summary>
        /// 專案經驗
        /// </summary>
        public List<ProjectExperience> ProjectExperienceList { get; set; }
        /// <summary>
        /// 面談結果
        /// </summary>
        public InterviewResults InterviewResults { get; set; }
    }

    /// <summary>
    /// 面談基本資料
    /// </summary>
    public class InterviewInfo
    {
        /// <summary>
        /// 應徵職位
        /// </summary>
        public string Vacancies { get; set; }
        /// <summary>
        /// 面談日期
        /// </summary>
        public string Interview_Date { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 性別
        /// </summary>
        public string Sex { get; set; }
        /// <summary>
        /// 生日
        /// </summary>
        public string Birthday { get; set; }
        /// <summary>
        /// 婚別
        /// </summary>
        public string Married { get; set; }
        /// <summary>
        /// e-mail
        /// </summary>
        public string Mail { get; set; }
        /// <summary>
        /// 地址
        /// </summary>
        public string Adress { get; set; }
        /// <summary>
        /// 手機
        /// </summary>
        public string CellPhone { get; set; }
        /// <summary>
        /// 圖片路徑
        /// </summary>
        public string Image { get; set; }
        /// <summary>
        /// 專長-程式語言
        /// </summary>
        public string Expertise_Language { get; set; }
        /// <summary>
        /// 專長-開發工具
        /// </summary>
        public string Expertise_Tools { get; set; }
        /// <summary>
        /// 專長-開發工具-Framwork
        /// </summary>
        public string Expertise_Tools_Framwork { get; set; }
        /// <summary>
        /// 專長-Devops
        /// </summary>
        public string Expertise_Devops { get; set; }
        /// <summary>
        /// 專長-作業系統
        /// </summary>
        public string Expertise_OS { get; set; }
        /// <summary>
        /// 專長-大數據
        /// </summary>
        public string Expertise_BigData { get; set; }
        /// <summary>
        /// 專長-資料庫
        /// </summary>
        public string Expertise_DataBase { get; set; }
        /// <summary>
        /// 專長-專業認證
        /// </summary>
        public string Expertise_Certification { get; set; }
        /// <summary>
        /// 最近三年內，是否有計畫繼續就學或出國深造?
        /// </summary>
        public string IsStudy { get; set; }
        /// <summary>
        /// 有無親友在本公司服務?
        /// </summary>
        public string IsService { get; set; }
        /// <summary>
        /// 有無親友在本公司服務?關係
        /// </summary>
        public string Relatives_Relationship { get; set; }
        /// <summary>
        /// 有無親友在本公司服務?姓名
        /// </summary>
        public string Relatives_Name { get; set; }
        /// <summary>
        /// 在工作中您最在乎什麼?
        /// </summary>
        public string Care_Work { get; set; }
        /// <summary>
        /// 希望待遇
        /// </summary>
        public string Hope_Salary { get; set; }
        /// <summary>
        /// 通知後多久可報到?
        /// </summary>
        public string When_Report { get; set; }
        /// <summary>
        /// 優點
        /// </summary>
        public string Advantage { get; set; }
        /// <summary>
        /// 缺點
        /// </summary>
        public string Disadvantages { get; set; }
        /// <summary>
        /// 嗜好
        /// </summary>
        public string Hobby { get; set; }
        /// <summary>
        /// 什麼原因吸引您來亦思應徵此工作? 您對亦思的第一印象是什麼?
        /// </summary>
        public string Attract_Reason { get; set; }
        /// <summary>
        /// 請描述您對未來的目標、希望。
        /// </summary>
        public string Future_Goal { get; set; }
        /// <summary>
        /// 您心目中的理想主管是什麼樣的人?
        /// </summary>
        public string Hope_Supervisor { get; set; }
        /// <summary>
        /// 您希望亦思給與您什麼承諾?(任何方面)
        /// </summary>
        public string Hope_Promise { get; set; }
        /// <summary>
        /// 覺得自己的哪個特質最強烈? 現在您有三十秒的自我推薦時間，您會透過什麼方式讓我們對您印象深刻?
        /// </summary>
        public string Introduction { get; set; }
        /// <summary>
        /// 服役期間
        /// </summary>
        public string During_Service { get; set; }
        /// <summary>
        /// 免疫說明
        /// </summary>
        public string Exemption_Reason { get; set; }
        /// <summary>
        /// 緊急通知人
        /// </summary>
        public string Urgent_Contact_Person { get; set; }
        /// <summary>
        /// 緊急通知人-關係
        /// </summary>
        public string Urgent_Relationship { get; set; }
        /// <summary>
        /// 緊急通知人-手機
        /// </summary>
        public string Urgent_CellPhone { get; set; }
        /// <summary>
        /// 學歷
        /// </summary>
        public string Education { get; set; }
        /// <summary>
        /// 語言能力
        /// </summary>
        public string Language { get; set; }
        /// <summary>
        /// 經歷
        /// </summary>
        public string Work_Experience { get; set; }
    }

    /// <summary>
    /// 面談結果(任用評定與備註)
    /// </summary>
    public class InterviewResult
    {
        /// <summary>
        /// 任用評定
        /// </summary>
        public string Appointment { get; set; }
        /// <summary>
        /// 面談結果-備註
        /// </summary>
        public string Results_Remark { get; set; }
    }

    /// <summary>
    /// 面談評語
    /// </summary>
    public class InterviewComments
    {
        /// <summary>
        /// 面談者
        /// </summary>
        public string Interviewer { get; set; }
        /// <summary>
        /// 面談結果
        /// </summary>
        public string Result { get; set; }
    }

    /// <summary>
    /// 面談結果(Excel)
    /// </summary>
    public class InterviewResults
    {
        /// <summary>
        /// 任用評定與備註
        /// </summary>
        public InterviewResult InterviewResult { get; set; }
        /// <summary>
        /// 面談評語
        /// </summary>
        public List<InterviewComments> InterviewCommentsList { get; set; }
    }

    /// <summary>
    /// 專案經驗
    /// </summary>
    public class ProjectExperience
    {
        /// <summary>
        /// 公司名稱
        /// </summary>
        public string Company { get; set; }
        /// <summary>
        /// 專案名稱
        /// </summary>
        public string Project_Name { get; set; }
        /// <summary>
        /// 作業系統
        /// </summary>
        public string OS { get; set; }
        /// <summary>
        /// 資料庫
        /// </summary>
        public string Database { get; set; }
        /// <summary>
        /// 職稱
        /// </summary>
        public string Position { get; set; }
        /// <summary>
        /// 程式語言
        /// </summary>
        public string Language { get; set; }
        /// <summary>
        /// 開發工具
        /// </summary>
        public string Tools { get; set; }
        /// <summary>
        /// 專案簡述
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// 起迄年月
        /// </summary>
        public string Start_End_Date { get; set; }
    }

    /// <summary>
    /// 附加檔案
    /// </summary>
    public class AttachedFiles
    {
        /// <summary>
        /// 檔案路徑
        /// </summary>
        public string File_Path { get; set; }
        /// <summary>
        /// 隸屬於人事資料或專案經驗
        /// </summary>
        public string Belong { get; set; }
    }
}
