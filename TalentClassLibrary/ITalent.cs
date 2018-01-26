using System.Collections.Generic;
using System.Data;

namespace TalentClassLibrary
{
    public interface ITalent
    {
        ////登入功能
        string SignIn(string account, string password);       
        bool ValidMemberStates(string states);
        string SignOut();
        ////權限管理功能
        DataTable SelectMemberInfoByAccount(string account);
        DataTable SelectMemberInfo();
        string AlertUpdatePassword(string account);
        string UpdateMemberStatesByAccount(string account, string states);
        string DelMemberByAccount(string account);
        string ValidInsertMember(string account);
        string InsertMember(string account);
        string UpdatePasswordByaccount(string account, string oldPassword, string newPassword, string checkNewPassword);
        string ValidNewPassword(string newPassword, string checkNewPassword);
        ////資料管理(查詢)
        List<string> SelectIdByFilter(string keyWord, string places, string expertises, string cooperationMode, string states,
                                      string startEditDate, string endEditDate, string IsInterview, string InterviewResult, 
                                      string startInterviewDate, string endInterviewDate);
        List<string> SelectIdByContact(string places, string expertises, string cooperationMode, string states);
        /*List<string> SelectIdByKeyWord(string keyWord);
        List<string> SelectIdByPlace(string place);
        List<string> SelectIdByexpertise(string expertise);*/
        DataTable SelectContactSituationInfoByCode(List<string> CodeList);
        string DelByCode(string Code);
        ////資料管理-聯絡狀況(編修)
        string ValidContactSituationData(DataTable validData);
        string ValidContactStatus(string contactStatus);
        string InsertContactSituation(DataTable inData);
        void DelContactSituation(string name);
        string ValidContactSituationInfoData(string name, string code, string sex, string mail, string cellPhone,
                                             string place, string expertise, string cooperationMode, string states);        
        string ValidCooperationMode(string cooperationMode);
        string ValidStates(string states);
        string InsertContactSituationInfoData(DataTable inData);
        string UpdateContactSituationInfoData(DataTable editData);
        ////資料管理-面談資料(查詢)
        DataTable SelectInterviewInfoByCode(List<string> CodeList);
        ////資料管理-面談資料(編修_個資)
        string ValidInterviewData(string vacancies, string name, string married, string interviewDate, string sex, string mail,
                                  string birthday, string cellPhone, string picture);
        string ValidMarried(string married);
        string ValidFilePath(string filePath);
        string InsertInterviewData(DataTable inData);
        string UpdateInterviewData(DataTable editData);
        ////資料管理-面談資料(編修_專案)
        string ValidProjectExperienceData(string company, string projectName, string os, string database, string position,
                                          string startingDate, string closingDate, string language, string tools, string description);
        string ValidProjectName(DataTable ProjectExperienceData);
        string InsertProjectExperience(DataTable inData);
        void DelProjectExperience(string name);
        ////匯出Excel 
        string ExportDataToExcel(DataTable data);
    }
}
