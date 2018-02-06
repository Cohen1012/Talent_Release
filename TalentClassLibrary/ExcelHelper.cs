using Newtonsoft.Json;
using ServiceStack;
using ShareClassLibrary;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.IO;
using TalentClassLibrary.Model;

namespace TalentClassLibrary
{
    /// <summary>
    /// 處理Excel的類別
    /// </summary>
    public class ExcelHelper
    {
        private static ExcelHelper excelHelper = new ExcelHelper();

        public static ExcelHelper GetInstance() => excelHelper;

        /// <summary>
        /// 紀錄大略發生的錯誤
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 匯出多筆聯繫狀況資料
        /// </summary>
        /// <param name="ContactSituationList"></param>
        public string ExportMultipleContactSituation(List<ContactSituation> ContactSituationList, string path)
        {
            try
            {
                //建立Workbook
                Workbook workbook = new Workbook();
                workbook.LoadTemplateFromFile(@"..\..\..\Template\TalentTemplate.xlsx");

                for (int i = 0; i < ContactSituationList.Count; i++)
                {
                    Worksheet sheet = workbook.CreateEmptySheet();
                    ////第一個Sheet當作Template
                    sheet.CopyFrom(workbook.Worksheets[0]);
                    ////Sheet命名
                    if (!string.IsNullOrEmpty(ContactSituationList[i].Info.Name))
                    {
                        sheet.Name = (i + 1) + "." + ContactSituationList[i].Info.Name;
                    }
                    else if (!string.IsNullOrEmpty(ContactSituationList[i].Code))
                    {
                        string[] code = ContactSituationList[i].Code.Split(new string[] { "\n" }, StringSplitOptions.None);
                        sheet.Name = (i + 1) + "." + code[0];
                    }

                    sheet = CreateContactSituationSheet(ContactSituationList, sheet, i);
                }

                ////隱藏Template Sheet
                for (int i = 0; i < 4; i++)
                {
                    workbook.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                }
                //儲存到物理路徑
                string strFullName = Path.Combine(path, "聯繫狀況" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                workbook.SaveToFile(strFullName, ExcelVersion.Version2010);
                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 匯出面談資料
        /// </summary>
        /// <param name="interviewDataList">面談資料</param>
        /// <param name="path">存檔路徑</param>
        /// <param name="count"></param>
        /// <returns></returns>
        public string ExportInterviewData(List<InterviewData> interviewDataList, string path, int count)
        {
            try
            {
                //建立Workbook
                Workbook workbook = new Workbook();
                workbook.LoadTemplateFromFile(@"..\..\..\Template\TalentTemplate.xlsx");
                for (int i = 0; i < interviewDataList.Count; i++)
                {
                    ////面談基本資訊Template
                    Worksheet sheet1 = workbook.CreateEmptySheet();
                    sheet1.Name = "人事資料" + (i + 1);
                    sheet1.CopyFrom(workbook.Worksheets[1]);
                    sheet1 = CreateInterviewInfoSheet(interviewDataList[i].InterviewInfo, sheet1);
                    ////專案經驗Template
                    Worksheet sheet2 = workbook.CreateEmptySheet();
                    sheet2.Name = "專案經驗" + (i + 1);
                    sheet2.CopyFrom(workbook.Worksheets[2]);
                    sheet2 = CreateProjectExperienceSheet(interviewDataList[i].ProjectExperienceList, sheet2);
                    ////面談結果Template
                    Worksheet sheet3 = workbook.CreateEmptySheet();
                    sheet3.Name = "面談結果" + (i + 1);
                    sheet3.CopyFrom(workbook.Worksheets[3]);
                    sheet3 = CreateInterviewResualtSheet(interviewDataList[i].InterviewResults, sheet3);
                }

                ////隱藏Template Sheet
                for (int i = 0; i < 4; i++)
                {
                    workbook.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                }
                //儲存到物理路徑
                string strFullName = Path.Combine(path, count + ".面談資料" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                workbook.SaveToFile(strFullName, ExcelVersion.Version2010);
                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 匯出所有資料
        /// </summary>
        /// <param name="contactSituationList">聯繫狀況資料</param>
        /// <param name="interviewDataList">面談資料</param>
        /// <param name="path">存檔路徑</param>
        /// <param name="count"></param>
        /// <returns></returns>
        public string ExportAllData(List<ContactSituation> contactSituationList, List<InterviewData> interviewDataList, string path, int count)
        {
            try
            {
                //建立Workbook
                Workbook workbook = new Workbook();
                workbook.LoadTemplateFromFile(@"..\..\..\Template\TalentTemplate.xlsx");
                //聯繫狀況Template
                Worksheet sheet = workbook.Worksheets[0];
                sheet = CreateContactSituationSheet(contactSituationList, sheet, 0);
                for (int i = 0; i < interviewDataList.Count; i++)
                {
                    ////面談基本資訊Template
                    Worksheet sheet1 = workbook.CreateEmptySheet();
                    sheet1.Name = "人事資料" + (i + 1);
                    sheet1.CopyFrom(workbook.Worksheets[1]);
                    sheet1 = CreateInterviewInfoSheet(interviewDataList[i].InterviewInfo, sheet1);
                    ////專案經驗Template
                    Worksheet sheet2 = workbook.CreateEmptySheet();
                    sheet2.Name = "專案經驗" + (i + 1);
                    sheet2.CopyFrom(workbook.Worksheets[2]);
                    sheet2 = CreateProjectExperienceSheet(interviewDataList[i].ProjectExperienceList, sheet2);
                    ////面談結果Template
                    Worksheet sheet3 = workbook.CreateEmptySheet();
                    sheet3.Name = "面談結果" + (i + 1);
                    sheet3.CopyFrom(workbook.Worksheets[3]);
                    sheet3 = CreateInterviewResualtSheet(interviewDataList[i].InterviewResults, sheet3);
                }

                ////隱藏Template Sheet
                for (int i = 1; i < 4; i++)
                {
                    workbook.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                }
                //儲存到物理路徑
                string strFullName = Path.Combine(path, count + ".聯繫狀況與面談資料" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                workbook.SaveToFile(strFullName, ExcelVersion.Version2010);
                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 匯入面談資料
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public DataSet ImportInterviewData(string path)
        {
            ErrorMessage = string.Empty;
            DataSet ds = new DataSet();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                ////面談基本資料
                Worksheet sheet = workbook.Worksheets["人事資料1"];
                ds.Tables.Add(this.ReadInterviewInfoSheet(sheet));
                ////面談結果
                Worksheet sheet2 = workbook.Worksheets["面談結果1"];
                ds.Tables.AddRange(this.ReadInterviewResultSheet(sheet2));
                ////專案經驗
                Worksheet sheet1 = workbook.Worksheets["專案經驗1"];
                ds.Tables.Add(this.ReadProjectExperienceSheet(sheet1));

                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "匯入失敗";
                ds.Clear();
                return ds;
            }
        }

        /// <summary>
        /// 匯入舊版資料
        /// </summary>
        /// <param name="path"></param>
        public List<ContactSituation> ImportOldTalent(string path)
        {
            ErrorMessage = string.Empty;
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dt = sheet.ExportDataTable();
                if (dt.Rows.Count == 0)
                {
                    ErrorMessage = "空的excel";
                    return contactSituationList;
                }

                foreach (DataRow dr in dt.Rows)
                {
                    ContactSituation contactSituation = new ContactSituation();
                    ContactInfo contactInfo = new ContactInfo();
                    ContactStatus contactStatus = new ContactStatus();
                    for (int i = 0; i < 10; i++)
                    {
                        switch (dt.Columns[i].ToString().Trim())
                        {
                            case "姓名":
                                contactInfo.Name = dr[i].ToString().Trim();
                                break;
                            case "性別":
                                contactInfo.Sex = dr[i].ToString().Trim();
                                break;
                            case "手機":
                                contactInfo.CellPhone = dr[i].ToString().Trim();
                                break;
                            case "郵件":
                                contactInfo.Mail = dr[i].ToString().Trim();
                                break;
                            case "學校":
                                if (!string.IsNullOrEmpty(dr[i].ToString().Trim()))
                                {
                                    contactStatus.Remarks += "學校\n" + dr[i].ToString().Trim() + "\n";
                                }

                                break;
                            case "地區":
                                contactInfo.Place = dr[i].ToString().Trim();
                                break;
                            case "最後編輯時間":
                                contactStatus.Contact_Date = dr[i].ToString().Trim();
                                break;
                            case "專長":
                                contactInfo.Skill = dr[i].ToString().Trim();
                                break;
                            case "互動":
                                if (!string.IsNullOrEmpty(dr[i].ToString().Trim()))
                                {
                                    contactStatus.Remarks += "互動\n" + dr[i].ToString().Trim() + "\n";
                                }

                                break;
                            case "評價":
                                if (!string.IsNullOrEmpty(dr[i].ToString().Trim()))
                                {
                                    contactStatus.Remarks += "評價\n" + dr[i].ToString().Trim() + "\n";
                                }

                                break;
                        }
                    }
                    contactSituation.Info = contactInfo;
                    contactSituation.Status = new List<ContactStatus> { contactStatus };
                    contactSituationList.Add(contactSituation);
                }

                return contactSituationList;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                contactSituationList.Clear();
                ErrorMessage = "讀取Excel發生錯誤";
                return contactSituationList;
            }
        }

        /// <summary>
        /// 匯入新版資料
        /// </summary>
        /// <param name="path"></param>
        public List<ContactSituation> ImportNewTalent(string path)
        {
            ErrorMessage = string.Empty;
            string msg = string.Empty;
            List<string> checkCodeIsRepeat = new List<string>(); ////檢查Excel內部的代碼是否有重複
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dt = sheet.ExportDataTable();
                if (dt.Rows.Count == 0)
                {
                    ErrorMessage = "空的excel";
                    return contactSituationList;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ContactSituation contactSituation = new ContactSituation();
                    ContactInfo contactInfo = new ContactInfo();
                    List<ContactStatus> contactStatusList = new List<ContactStatus>();
                    for (int j = 0; j < 29; j++)
                    {
                        ////先不處理聯繫狀況
                        if (j == 3 || j == 4 || j == 5)
                        {
                            continue;
                        }

                        switch (dt.Columns[j].ToString().Trim())
                        {
                            case "姓名":
                                contactInfo.Name = dt.Rows[i].ItemArray[j].ToString().Trim();
                                break;
                            case "地點":
                                contactInfo.Place = dt.Rows[i].ItemArray[j].ToString().Trim();
                                break;
                            case "1111/104代碼":
                                ////檢查Excel內部的代碼是否有重複
                                string[] codeList = dt.Rows[i].ItemArray[j].ToString().Trim().Split('\n');
                                foreach (string code in codeList)
                                {
                                    if (checkCodeIsRepeat.Contains(code))
                                    {
                                        contactSituationList.Clear();
                                        ErrorMessage = "第" + (i + 1) + "行" + code + "重複\n請檢查Excel";
                                        return contactSituationList;
                                    }

                                    checkCodeIsRepeat.Add(code);
                                }

                                contactSituation.Code = dt.Rows[i].ItemArray[j].ToString().Trim();
                                break;
                            case "JAVA":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "JSP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Android APP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "ASP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "C/C++":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "C#":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "ASP.NET":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "VB.NET":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "VB6":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "HTML":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Javascript":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Bootstrap":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Delphi":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "PHP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "研替":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Hadoop":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "ETL":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "R":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "notes":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "UI/UX":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "資料庫":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Linux":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                        }
                    }

                    ////聯繫狀況
                    do
                    {
                        ContactStatus contactStatus = new ContactStatus();
                        for (int z = 3; z <= 5; z++)
                        {
                            switch (dt.Columns[z].ToString().Trim())
                            {
                                case "日期":
                                    msg = Valid.GetInstance().ValidDateFormat(dt.Rows[i].ItemArray[z].ToString().Trim());
                                    if (msg != string.Empty)
                                    {
                                        contactSituationList.Clear();
                                        ErrorMessage = "第" + (i + 1) + "行" + msg + "\n請檢查Excel";
                                        return contactSituationList;
                                    }

                                    contactStatus.Contact_Date = dt.Rows[i].ItemArray[z].ToString().Trim();
                                    break;
                                case "聯絡狀況":
                                    if (string.IsNullOrEmpty(dt.Rows[i].ItemArray[z].ToString().Trim()))
                                    {
                                        contactStatus.Contact_Status = "(無)";
                                    }
                                    else
                                    {
                                        msg = Talent.GetInstance().ValidContactStatus(dt.Rows[i].ItemArray[z].ToString().Trim());
                                        if (msg != string.Empty)
                                        {
                                            contactSituationList.Clear();
                                            ErrorMessage = "第" + (i + 1) + "行" + msg + "\n請檢查Excel";
                                            return contactSituationList;
                                        }

                                        contactStatus.Contact_Status = dt.Rows[i].ItemArray[z].ToString().Trim();
                                    }
                                    break;
                                case "說明":
                                    contactStatus.Remarks = dt.Rows[i].ItemArray[z].ToString().Trim();
                                    break;
                            }
                        }

                        contactStatusList.Add(contactStatus);

                        i++;
                        if (i >= dt.Rows.Count)
                        {
                            break;
                        }
                    } while (string.IsNullOrEmpty(dt.Rows[i].ItemArray[0].ToString().Trim()) && string.IsNullOrEmpty(dt.Rows[i].ItemArray[1].ToString().Trim()));

                    i--; ////因為在do迴圈有先i++判斷下一列的姓名與代碼是否有值，因此在跳出迴圈時，要把它減回來
                    contactInfo.Skill = contactInfo.Skill.RemoveEndWithDelimiter(",");
                    contactSituation.Info = contactInfo;
                    contactSituation.Status = contactStatusList;
                    contactSituationList.Add(contactSituation);
                }

                msg = Talent.GetInstance().ValidCodeIsRepeat(checkCodeIsRepeat);
                if (msg != string.Empty)
                {
                    contactSituationList.Clear();
                    ErrorMessage = msg + "\n請檢查Excel";
                    return contactSituationList;
                }

                return contactSituationList;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                contactSituationList.Clear();
                ErrorMessage = "讀取Excel發生錯誤";
                return contactSituationList;
            }
        }

        /// <summary>
        /// 讀取面談結果的Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable[] ReadInterviewResultSheet(Worksheet sheet)
        {
            DataTable[] dt = new DataTable[2];
            int row = 2; ////紀錄目前在哪一行
            InterviewResult interviewResult = new InterviewResult();
            List<InterviewComments> interviewCommentList = new List<InterviewComments>();
            if (sheet.Range["B" + row].Value.Trim() != "任用評定")
            {
                throw new Exception("面談結果sheet格式不符合");
            }

            row = 5;
            interviewResult.Appointment = string.Empty;
            ////任用評定
            for (int i = 5; i <= 7; i++)
            {
                if (sheet.Range["B" + row].Value.Trim().StartsWith("■"))
                {
                    interviewResult.Appointment = sheet.Range["B" + row].Value.Trim().RemoveStartsWithDelimiter("■");
                    break;
                }
                row++;
            }

            row = 13;
            ////面談評語
            while (sheet.Range["B" + row].Value.Trim() != "備註")
            {
                InterviewComments interviewComments = new InterviewComments
                {
                    Interviewer = sheet.Range["C" + row].Value.Trim(),
                    Result = sheet.Range["G" + row].Value.Trim()
                };

                if (!interviewComments.TResultIsEmtpty())
                {
                    interviewCommentList.Add(interviewComments);
                }
                row++;
                if (sheet.Range["B" + row].Value.Trim() == "備註")
                {
                    break;
                }

                row++;
            }

            row += 3;
            interviewResult.Results_Remark = sheet.Range["B" + row].Value.Trim();
            List<InterviewResult> interviewResultList = new List<InterviewResult>
            {
                interviewResult
            };

            if(interviewCommentList.Count == 0)
            {
                interviewCommentList.Add(new InterviewComments());
            }

            dt[1] = interviewResultList.ListToDataTable();
            dt[0] = interviewCommentList.ListToDataTable();

            return dt;
        }

        /// <summary>
        /// 讀取專案經驗的Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable ReadProjectExperienceSheet(Worksheet sheet)
        {
            int row = 2; ////紀錄目前在哪一行
            List<ProjectExperience> projectExperienceList = new List<ProjectExperience>();
            while (!string.IsNullOrEmpty(sheet.Range["B" + row].Value.Trim()) && sheet.Range["B" + row].Value.Trim().Equals("公司名稱"))
            {
                ProjectExperience projectExperience = new ProjectExperience
                {
                    Company = sheet.Range["C" + row].Value.Trim(),
                    Position = sheet.Range["H" + row].Value.Trim(),
                };

                row++;
                projectExperience.Project_Name = sheet.Range["C" + row].Value.Trim();
                projectExperience.Start_End_Date = sheet.Range["H" + row].Value.Trim();
                row++;
                projectExperience.OS = sheet.Range["C" + row].Value.Trim();
                projectExperience.Language = sheet.Range["H" + row].Value.Trim();
                row++;
                projectExperience.Database = sheet.Range["C" + row].Value.Trim();
                projectExperience.Tools = sheet.Range["H" + row].Value.Trim();
                row++;
                projectExperience.Description = sheet.Range["C" + row].Value.Trim();
                row += 3;
                projectExperienceList.Add(projectExperience);
            }

            if(projectExperienceList.Count == 0)
            {
                projectExperienceList.Add(new ProjectExperience());
            }

            return projectExperienceList.ListToDataTable();
        }

        /// <summary>
        /// 讀取面談基本資料的Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable ReadInterviewInfoSheet(Worksheet sheet)
        {
            int row = 0; ////紀錄目前在哪一行
            DataTable dt = new DataTable();
            if (sheet.Range["E2"].Value.Trim() != "應徵職缺")
            {
                throw new Exception("面談結果sheet格式不符合");
            }

            InterviewInfo interviewInfo = new InterviewInfo
            {
                Vacancies = sheet.Range["F2"].Value.Trim(),
                Name = sheet.Range["F4"].Value.Trim(),
                Married = sheet.Range["F6"].Value.Trim(),
                Adress = sheet.Range["F8"].Value.Trim(),
                Urgent_Contact_Person = sheet.Range["F10"].Value.Trim(),
                Interview_Date = sheet.Range["I2"].Value.Trim(),
                Sex = sheet.Range["I4"].Value.Trim(),
                Mail = sheet.Range["I6"].Value.Trim(),
                Urgent_Relationship = sheet.Range["I10"].Value.Trim(),
                Birthday = sheet.Range["L4"].Value.Trim(),
                CellPhone = sheet.Range["L6"].Value.Trim(),
                Urgent_CellPhone = sheet.Range["L10"].Value.Trim(),
            };

            ////處理圖片
            if (sheet.Pictures.Count > 0)
            {
                ExcelPicture picture = sheet.Pictures[0];
                string a = Path.GetTempFileName();
                picture.Picture.Save(a, ImageFormat.Png);
                interviewInfo.Image = a;
                //picture.Picture.Save(@"..\..\..\Template\image.png", ImageFormat.Png);
            }

            row = 13;
            ////教育程度
            List<Education> educationList = new List<Education>();
            while (sheet.Range["C" + row].Value.Trim() != "經歷")
            {
                Education education = new Education
                {
                    School = sheet.Range["D" + row].Value.Trim(),
                    Department = sheet.Range["E" + row].Value.Trim(),
                    Start_End_Date = sheet.Range["F" + row].Value.Trim(),
                    Is_Graduation = sheet.Range["G" + row].Value.Trim(),
                    Remark = sheet.Range["H" + row].Value.Trim(),
                };

                if (!education.TResultIsEmtpty())
                {
                    educationList.Add(education);
                }
                row++;
            }
            interviewInfo.Education = JsonConvert.SerializeObject(educationList, Formatting.Indented);
            interviewInfo.Education = interviewInfo.Education == "[]" ? string.Empty : interviewInfo.Education;

            row++;
            ////工作經驗
            List<WorkExperience> workExperienceList = new List<WorkExperience>();
            while (sheet.Range["C" + row].Value.Trim() != "兵役")
            {
                WorkExperience workExperience = new WorkExperience
                {
                    Institution_name = sheet.Range["D" + row].Value.Trim(),
                    Position = sheet.Range["E" + row].Value.Trim(),
                    Start_End_Date = sheet.Range["F" + row].Value.Trim(),
                    Start_Salary = sheet.Range["G" + row].Value.Trim(),
                    End_Salary = sheet.Range["H" + row].Value.Trim(),
                    Leaving_Reason = sheet.Range["I" + row].Value.Trim(),
                };

                if (!workExperience.TResultIsEmtpty())
                {
                    workExperienceList.Add(workExperience);
                }
                row++;
            }
            interviewInfo.Work_Experience = JsonConvert.SerializeObject(workExperienceList, Formatting.Indented);
            interviewInfo.Work_Experience = interviewInfo.Work_Experience == "[]" ? string.Empty : interviewInfo.Work_Experience;

            ////兵役
            row++;
            interviewInfo.IsService = sheet.Range["D" + row].Value.Trim();
            interviewInfo.Exemption_Reason = sheet.Range["E" + row].Value.Trim();
            row += 3;
            ////專長
            ////專長-程式語言
            interviewInfo.Expertise_Language = GetExpertise(sheet, row, 11, 13);
            row += 2;
            ////專長-開發工具
            interviewInfo.Expertise_Tools = GetExpertise(sheet, row, 7, 12);
            interviewInfo.Expertise_Tools_Framwork = sheet.Range[row, 9].Value.Trim();
            row += 2;
            ////專長-Devops
            interviewInfo.Expertise_Devops = GetExpertise(sheet, row, 9, 11);
            row += 2;
            ////專長-作業系統
            interviewInfo.Expertise_OS = GetExpertise(sheet, row, 9, 11);
            row += 2;
            ////專長-大數據
            interviewInfo.Expertise_BigData = GetExpertise(sheet, row, 7, 9);
            row += 2;
            ////專長-資料庫
            interviewInfo.Expertise_DataBase = GetExpertise(sheet, row, 8, 10);
            row += 2;
            ////專長-專業認證
            interviewInfo.Expertise_Certification = GetExpertise(sheet, row, 9, 11);
            row += 3;
            ////語言能力
            List<Language> languageList = new List<Language>();
            while (sheet.Range["C" + row].Value.Trim() != "最近三年內，是否有計畫繼續就學或出國深造?")
            {
                Language language = new Language
                {
                    Language_Name = sheet.Range["D" + row].Value.Trim(),
                    Listen = sheet.Range["E" + row].Value.Trim(),
                    Speak = sheet.Range["F" + row].Value.Trim(),
                    Read = sheet.Range["G" + row].Value.Trim(),
                    Write = sheet.Range["H" + row].Value.Trim(),
                };

                if (!language.TResultIsEmtpty())
                {
                    languageList.Add(language);
                }
                row++;
            }
            interviewInfo.Language = JsonConvert.SerializeObject(languageList, Formatting.Indented);
            interviewInfo.Language = interviewInfo.Language == "[]" ? string.Empty : interviewInfo.Language;
            row++;
            interviewInfo.IsStudy = sheet.Range["C" + row].Value.Trim();
            row += 2;
            interviewInfo.IsService = sheet.Range["F" + row].Value.Trim();
            interviewInfo.Relatives_Relationship = sheet.Range["J" + row].Value.Trim();
            interviewInfo.Relatives_Name = sheet.Range["N" + row].Value.Trim();
            row += 2;
            interviewInfo.Care_Work = sheet.Range["F" + row].Value.Trim();
            interviewInfo.Hope_Salary = sheet.Range["J" + row].Value.Trim();
            interviewInfo.When_Report = sheet.Range["N" + row].Value.Trim();
            row += 2;
            interviewInfo.Advantage = sheet.Range["D" + row].Value.Trim();
            interviewInfo.Disadvantages = sheet.Range["J" + row].Value.Trim();
            row += 2;
            interviewInfo.Hobby = sheet.Range["D" + row].Value.Trim();
            row += 3;
            interviewInfo.Attract_Reason = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Future_Goal = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Hope_Supervisor = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Hope_Promise = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Introduction = sheet.Range["C" + row].Value.Trim();

            List<InterviewInfo> interviewInfoList = new List<InterviewInfo>
            {
                interviewInfo
            };

            dt = interviewInfoList.ListToDataTable();

            return dt;

        }

        /// <summary>
        /// 面談資本資料的Sheet
        /// </summary>
        /// <param name="interviewInfo">面談基本資訊</param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Worksheet CreateInterviewInfoSheet(InterviewInfo interviewInfo, Worksheet sheet)
        {
            int row = 0; ////紀錄目前在哪一行
            ////照片
            if (!string.IsNullOrEmpty(interviewInfo.Image))
            {
                string picPath = interviewInfo.Image;
                ExcelPicture picture = sheet.Pictures.Add(2, 2, picPath);
                picture.Height = 135;
                picture.Width = 145;
            }

            ////面談基本資料
            sheet.Range["F2"].Text = interviewInfo.Vacancies;
            sheet.Range["F4"].Text = interviewInfo.Name;
            sheet.Range["F6"].Text = interviewInfo.Married;
            sheet.Range["F8"].Text = interviewInfo.Adress;
            sheet.Range["F10"].Text = interviewInfo.Urgent_Contact_Person;
            sheet.Range["I2"].Text = interviewInfo.Interview_Date;
            sheet.Range["I4"].Text = interviewInfo.Sex;
            sheet.Range["I6"].Text = interviewInfo.Mail;
            sheet.Range["I10"].Text = interviewInfo.Urgent_Relationship;
            sheet.Range["L4"].Text = string.IsNullOrEmpty(interviewInfo.Birthday) ? string.Empty : interviewInfo.Birthday;
            sheet.Range["L6"].Text = interviewInfo.CellPhone;
            sheet.Range["L10"].Text = interviewInfo.Urgent_CellPhone;
            ////學歷
            List<Education> educationList = new List<Education>();
            if (!string.IsNullOrEmpty(interviewInfo.Education))
            {
                educationList = interviewInfo.Education.FromJson<List<Education>>();
            }

            row = 12;
            for (int i = 0; i < educationList.Count; i++)
            {
                row++;
                sheet.Range[row, 4].Text = string.IsNullOrEmpty(educationList[i].School) ? string.Empty : educationList[i].School;
                sheet.Range[row, 5].Text = string.IsNullOrEmpty(educationList[i].Department) ? string.Empty : educationList[i].Department;
                sheet.Range[row, 6].Text = string.IsNullOrEmpty(educationList[i].Start_End_Date) ? string.Empty : educationList[i].Start_End_Date;
                sheet.Range[row, 7].Text = string.IsNullOrEmpty(educationList[i].Is_Graduation) ? string.Empty : educationList[i].Is_Graduation;
                sheet.Range[row, 8].Text = string.IsNullOrEmpty(educationList[i].Remark) ? string.Empty : educationList[i].Remark;
                sheet.Range["H" + row + ":M" + row].Merge();
                sheet.Range["D" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
            }
            sheet.Range["D12:M" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D12:M" + row].BorderInside(LineStyleType.Thin);
            /////工作經驗
            List<WorkExperience> workExperienceList = new List<WorkExperience>();
            if (!string.IsNullOrEmpty(interviewInfo.Work_Experience))
            {
                workExperienceList = interviewInfo.Work_Experience.FromJson<List<WorkExperience>>();
            }

            row += 2;
            sheet.Range[row, 3].Text = "經歷";
            sheet.Range[row, 4].Text = "機構名稱";
            sheet.Range[row, 5].Text = "職稱";
            sheet.Range[row, 6].Text = "起迄年月";
            sheet.Range[row, 7].Text = "到職薪資";
            sheet.Range[row, 8].Text = "離職薪資";
            sheet.Range[row, 9].Text = "離職原因";
            sheet.Range["I" + row + ":M" + row].Merge();
            sheet.Range["D" + row + ":I" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;

            for (int i = 0; i < workExperienceList.Count; i++)
            {
                row++;
                sheet.Range[row, 4].Text = string.IsNullOrEmpty(workExperienceList[i].Institution_name) ? string.Empty : workExperienceList[i].Institution_name;
                sheet.Range[row, 5].Text = string.IsNullOrEmpty(workExperienceList[i].Position) ? string.Empty : workExperienceList[i].Position;
                sheet.Range[row, 6].Text = string.IsNullOrEmpty(workExperienceList[i].Start_End_Date) ? string.Empty : workExperienceList[i].Start_End_Date;
                sheet.Range[row, 7].Text = string.IsNullOrEmpty(workExperienceList[i].Start_Salary) ? string.Empty : workExperienceList[i].Start_Salary;
                sheet.Range[row, 8].Text = string.IsNullOrEmpty(workExperienceList[i].End_Salary) ? string.Empty : workExperienceList[i].End_Salary;
                sheet.Range[row, 9].Text = string.IsNullOrEmpty(workExperienceList[i].Leaving_Reason) ? string.Empty : workExperienceList[i].Leaving_Reason;
                sheet.Range["I" + row + ":M" + row].Merge();
                sheet.Range["D" + row + ":M" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["G" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Right;
            }
            sheet.Range["D" + (row - workExperienceList.Count) + ":M" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D" + (row - workExperienceList.Count) + ":M" + row].BorderInside(LineStyleType.Thin);
            ////兵役
            row += 2;
            sheet.Range[row, 3].Text = "兵役";
            sheet.Range[row, 4].Text = "服役期間";
            sheet.Range[row, 5].Text = "免役說明";
            sheet.Range["E" + row + ":M" + row].Merge();
            sheet.Range["D" + row + ":E" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;
            row++;
            sheet.Range[row, 4].Text = interviewInfo.During_Service;
            sheet.Range[row, 5].Text = interviewInfo.Exemption_Reason;
            sheet.Range["E" + row + ":M" + row].Merge();
            sheet.Range["D" + row + ":M" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["D" + (row - 1) + ":M" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D" + (row - 1) + ":M" + row].BorderInside(LineStyleType.Thin);
            ////專長
            row += 2;
            sheet.Range[row, 3].Text = "專長";
            sheet.Range[row, 4].Text = "※請盡可能寫出您所具備的專長及與工作有關的技能";
            sheet.Range["D" + row + ":G" + row].Merge();
            sheet.Range["D" + row + ":G" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            ////----程式語言
            sheet.Range[row, 4].Text = "程式語言";
            string[] Expertise_Language = { "C/C++", "Java", "C#", "VB", "Scala", "R", "Python" };
            for (int i = 5; i <= 11; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Language[i - 5];
            }

            sheet.Range[row, 12].Text = "其他";
            sheet.Range["M" + row + ":O" + row].Merge();
            sheet.Range["M" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 13].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Language, sheet, row, Expertise_Language, 13);
            ////----開發工具
            row += 2;
            sheet.Range[row, 4].Text = "開發工具";
            string[] Expertise_Tools = { "Visual Studio", "Eclipse", "Xcode" };
            for (int i = 5; i <= 7; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Tools[i - 5];
            }

            sheet.Range[row, 8].Text = "Framwork";
            sheet.Range["I" + row + ":J" + row].Merge();
            sheet.Range["I" + row + ":J" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 9].Text = interviewInfo.Expertise_Tools_Framwork ?? string.Empty;
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 11].Text = "其他";
            sheet.Range["L" + row + ":O" + row].Merge();
            sheet.Range["L" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 12].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Tools, sheet, row, Expertise_Tools, 12);
            ////----Devops
            row += 2;
            sheet.Range[row, 4].Text = "Devops";
            string[] Expertise_Devops = { "Docker", "Git", "Vagrant", "Puppet", "TFS" };
            for (int i = 5; i <= 9; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Devops[i - 5];
            }

            sheet.Range[row, 10].Text = "其他";
            sheet.Range["K" + row + ":O" + row].Merge();
            sheet.Range["K" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 11].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Devops, sheet, row, Expertise_Devops, 11);
            ////----作業系統
            row += 2;
            sheet.Range[row, 4].Text = "作業系統";
            string[] Expertise_OS = { "Windows", "Linux", "Unix", "Android", "IOS" };
            for (int i = 5; i <= 9; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_OS[i - 5];
            }

            sheet.Range[row, 10].Text = "其他";
            sheet.Range["K" + row + ":O" + row].Merge();
            sheet.Range["K" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 11].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_OS, sheet, row, Expertise_OS, 11);
            ////----大數據
            row += 2;
            sheet.Range[row, 4].Text = "大數據";
            string[] Expertise_BigData = { "Hadoop", "Spark", "Cloudera" };
            for (int i = 5; i <= 7; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_BigData[i - 5];
            }

            sheet.Range[row, 8].Text = "其他";
            sheet.Range["I" + row + ":O" + row].Merge();
            sheet.Range["I" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_BigData, sheet, row, Expertise_BigData, 9);
            ////----資料庫
            row += 2;
            sheet.Range[row, 4].Text = "資料庫";
            string[] Expertise_DataBase = { "Oracle", "My SQL", "MS SQL", "Hbase" };
            for (int i = 5; i <= 8; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_DataBase[i - 5];
            }

            sheet.Range[row, 9].Text = "其他";
            sheet.Range["J" + row + ":O" + row].Merge();
            sheet.Range["J" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_DataBase, sheet, row, Expertise_DataBase, 10);
            ////----專業認證
            row += 2;
            sheet.Range[row, 4].Text = "專業認證";
            string[] Expertise_Certification = { "SCJP", "SCWCD", "MCAD", "MCTS", "PMP" };
            for (int i = 5; i <= 9; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Certification[i - 5];
            }

            sheet.Range[row, 10].Text = "其他";
            sheet.Range["K" + row + ":O" + row].Merge();
            sheet.Range["K" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 11].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Certification, sheet, row, Expertise_Certification, 11);
            /////語言能力
            List<Language> languageList = new List<Language>();
            if (!string.IsNullOrEmpty(interviewInfo.Language))
            {
                languageList = interviewInfo.Language.FromJson<List<Language>>();
            }

            row += 2;
            sheet.Range[row, 3].Text = "語言能力";
            sheet.Range[row, 4].Text = "語言";
            sheet.Range[row, 5].Text = "聽";
            sheet.Range[row, 6].Text = "說";
            sheet.Range[row, 7].Text = "讀";
            sheet.Range[row, 8].Text = "寫";
            sheet.Range["D" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;

            for (int i = 0; i < languageList.Count; i++)
            {
                row++;
                sheet.Range[row, 4].Text = string.IsNullOrEmpty(languageList[i].Language_Name) ? string.Empty : languageList[i].Language_Name;
                sheet.Range[row, 5].Text = string.IsNullOrEmpty(languageList[i].Listen) ? string.Empty : languageList[i].Listen;
                sheet.Range[row, 6].Text = string.IsNullOrEmpty(languageList[i].Speak) ? string.Empty : languageList[i].Speak;
                sheet.Range[row, 7].Text = string.IsNullOrEmpty(languageList[i].Read) ? string.Empty : languageList[i].Read;
                sheet.Range[row, 8].Text = string.IsNullOrEmpty(languageList[i].Write) ? string.Empty : languageList[i].Write;
                sheet.Range["D" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;
            }
            sheet.Range["D" + (row - languageList.Count) + ":H" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D" + (row - languageList.Count) + ":H" + row].BorderInside(LineStyleType.Thin);
            ////問題
            row += 2;
            sheet.Range[row, 3].Text = "最近三年內，是否有計畫繼續就學或出國深造?";
            sheet.Range["C" + row + ":F" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.IsStudy;
            row += 2;
            sheet.Range[row, 3].Text = "有無親友在本公司服務?";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["F" + row + ":G" + row].Merge();
            sheet.Range["F" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 6].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 6].Text = interviewInfo.IsService;
            sheet.Range[row, 9].Text = "關係";
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["J" + row + ":K" + row].Merge();
            sheet.Range["J" + row + ":K" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 10].Text = interviewInfo.Relatives_Relationship;
            sheet.Range[row, 13].Text = "姓名";
            sheet.Range[row, 13].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["N" + row + ":O" + row].Merge();
            sheet.Range["N" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 14].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 14].Text = interviewInfo.Relatives_Name;
            row += 2;
            sheet.Range[row, 3].Text = "在工作中您最在乎什麼?";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["F" + row + ":G" + row].Merge();
            sheet.Range["F" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 6].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 6].Text = interviewInfo.Care_Work;
            sheet.Range[row, 9].Text = "希望待遇";
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["J" + row + ":K" + row].Merge();
            sheet.Range["J" + row + ":K" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 10].Text = interviewInfo.Hope_Salary;
            sheet.Range[row, 13].Text = "通知後多久可報到?";
            sheet.Range[row, 13].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["N" + row + ":O" + row].Merge();
            sheet.Range["N" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 14].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 14].Text = interviewInfo.When_Report;
            row += 2;
            sheet.Range[row, 3].Text = "優點";
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["D" + row + ":G" + row].Merge();
            sheet.Range["D" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 4].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 4].Text = interviewInfo.Advantage;
            sheet.Range[row, 9].Text = "缺點";
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["J" + row + ":O" + row].Merge();
            sheet.Range["J" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 10].Text = interviewInfo.Disadvantages;
            row += 2;
            sheet.Range[row, 3].Text = "嗜好";
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["D" + row + ":G" + row].Merge();
            sheet.Range["D" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 4].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 4].Text = interviewInfo.Hobby;
            row += 2;
            sheet.Range[row, 3].Text = "什麼原因吸引您來亦思應徵此工作? 您對亦思的第一印象是什麼?";
            sheet.Range["C" + row + ":G" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Attract_Reason;
            row += 2;
            sheet.Range[row, 3].Text = "請描述您對未來的目標、希望。";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Future_Goal;
            row += 2;
            sheet.Range[row, 3].Text = "您心目中的理想主管是什麼樣的人?";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Hope_Supervisor;
            row += 2;
            sheet.Range[row, 3].Text = "您希望亦思給與您什麼承諾?(任何方面)";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Hope_Promise;
            row += 2;
            sheet.Range[row, 3].Text = "覺得自己的哪個特質最強烈? 現在您有三十秒的自我推薦時間，您會透過什麼方式讓我們對您印象深刻?";
            sheet.Range["C" + row + ":K" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Introduction;
            return sheet;
        }

        /// <summary>
        /// 打專長組成字串
        /// </summary>
        /// <param name="row">第幾行</param>
        /// <param name="endExpertiseIndex">最後預設的專長的欄數</param>
        /// <param name="otherExpertiseIndex">其他專長的欄數</param>
        /// <returns></returns>
        private string GetExpertise(Worksheet sheet, int row, int endExpertiseIndex, int otherExpertiseIndex)
        {
            string expertise = string.Empty;
            for (int i = 5; i <= endExpertiseIndex; i++)
            {
                if (sheet.Range[row, i].Value.Trim().StartsWith("■"))
                {
                    expertise += sheet.Range[row, i].Value.Trim().RemoveStartsWithDelimiter("■") + ",";
                }
            }

            expertise += sheet.Range[row, otherExpertiseIndex].Value.Trim();
            expertise = expertise.RemoveEndWithDelimiter(",");

            return expertise;
        }

        /// <summary>
        /// 勾選專長
        /// </summary>
        /// <param name="Expertise">專長標題</param>
        /// <param name="sheet"></param>
        /// <param name="row">第幾列</param>
        /// <param name="DefaultExpertise">該專長的預設項目</param>
        /// <param name="column">"其他"欄位所在的Column</param>
        /// <returns></returns>
        private Worksheet ChcekedExpertise(string Expertise, Worksheet sheet, int row, string[] DefaultExpertise, int column)
        {
            if (!string.IsNullOrEmpty(Expertise))
            {
                string[] expertiseDevops = Expertise.Split(',');
                for (int i = 0; i < expertiseDevops.Length; i++)
                {
                    for (int j = 0; j < DefaultExpertise.Length; j++)
                    {
                        if (DefaultExpertise[j] == expertiseDevops[i])
                        {
                            sheet.Range[row, (j + 5)].Text = "■" + DefaultExpertise[j];
                            break;
                        }

                        if (j == DefaultExpertise.Length - 1)
                        {
                            sheet.Range[row, column].Text += expertiseDevops[i] + ",";
                        }
                    }
                }

                sheet.Range[row, column].Text = this.RemoveEndWithComma(sheet.Range[row, column].Text);
            }

            return sheet;
        }

        /// <summary>
        /// 專案經驗的Sheet
        /// </summary>
        /// <param name="projectExperienceList">專案經驗資料</param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Worksheet CreateProjectExperienceSheet(List<ProjectExperience> projectExperienceList, Worksheet sheet)
        {
            int rowCount = 2; ////紀錄目前所在的列數
            for (int i = 0; i < projectExperienceList.Count; i++)
            {
                sheet.Range[rowCount, 2].Text = "公司名稱";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Company;
                sheet.Range[rowCount, 7].Text = "職稱";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Position;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "專案名稱";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Project_Name;
                sheet.Range[rowCount, 7].Text = "起迄年月";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Start_End_Date;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "作業系統";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].OS;
                sheet.Range[rowCount, 7].Text = "程式語言";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Language;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "資料庫";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Database;
                sheet.Range[rowCount, 7].Text = "開發工具";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Tools;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "專案簡述";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":L" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":L" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Description;
                rowCount++;
                sheet.Range["B" + (rowCount - 5) + ":M" + rowCount].BorderAround(LineStyleType.Thick);
                rowCount += 2;
            }
            return sheet;
        }

        /// <summary>
        /// 面談結果
        /// </summary>
        /// <param name="interviewResults"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Worksheet CreateInterviewResualtSheet(InterviewResults interviewResults, Worksheet sheet)
        {
            int rowCount = 11;
            ////任用評定
            InterviewResult interviewResult = interviewResults.InterviewResult;
            switch (interviewResult.Appointment)
            {
                case "錄用":
                    sheet.Range["B5"].Text = "■" + interviewResult.Appointment;
                    break;
                case "不錄用":
                    sheet.Range["B6"].Text = "■" + interviewResult.Appointment;
                    break;
                case "暫保留":
                    sheet.Range["B7"].Text = "■" + interviewResult.Appointment;
                    break;
            }
            ////面談評語
            List<InterviewComments> interviewCommentsList = interviewResults.InterviewCommentsList;
            if (interviewCommentsList.Count == 0)
            {
                rowCount += 2;
            }

            for (int i = 0; i < interviewCommentsList.Count; i++)
            {
                rowCount += 2;
                sheet.Range["B" + rowCount].Text = "面談者";
                sheet.Range["B" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":D" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":D" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range["C" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount].Text = interviewCommentsList[i].Interviewer;
                sheet.Range["F" + rowCount].Text = "面談結果";
                sheet.Range["F" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["G" + rowCount + ":L" + rowCount].Merge();
                sheet.Range["G" + rowCount + ":L" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range["G" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["G" + rowCount].Text = interviewCommentsList[i].Result;
            }
            rowCount++;
            sheet.Range["B" + (rowCount - interviewCommentsList.Count - 2) + ":M" + rowCount].BorderAround(LineStyleType.Medium);
            rowCount += 2;
            sheet.Range["B" + rowCount + ":C" + (rowCount + 1)].Merge();
            sheet.Range["B" + rowCount].Text = "備註";
            sheet.Range["B" + rowCount + ":C" + (rowCount + 1)].BorderAround(LineStyleType.Medium);
            rowCount += 3;
            sheet.Range["B" + rowCount + ":M" + (rowCount + 2)].Merge();
            sheet.Range["B" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["B" + rowCount].Text = string.IsNullOrEmpty(interviewResult.Results_Remark) ? string.Empty : interviewResult.Results_Remark;
            sheet.Range["B" + rowCount + ":M" + (rowCount + 2)].BorderAround(LineStyleType.Medium);


            return sheet;
        }
        /// <summary>
        /// 聯繫狀況的Sheet
        /// </summary>
        /// <param name="ContactSituationList">聯繫狀況資料</param>
        /// <param name="sheet"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private Worksheet CreateContactSituationSheet(List<ContactSituation> ContactSituationList, Worksheet sheet, int i)
        {
            sheet.Range["C2"].Text = ContactSituationList[i].Info.Name;
            sheet.Range["C2"].Style.Font.Underline = FontUnderlineType.None;
            sheet.Range["C4"].Text = ContactSituationList[i].Code;
            sheet.Range["C4"].Style.WrapText = true;
            sheet.Range["C6"].Text = ContactSituationList[i].Info.Sex;
            sheet.Range["C8"].Text = ContactSituationList[i].Info.Mail;
            sheet.Range["C10"].Text = ContactSituationList[i].Info.CellPhone;

            sheet.Range["J2"].Text = ContactSituationList[i].Info.Place;
            sheet.Range["J4"].Text = ContactSituationList[i].Info.Skill;
            sheet.Range["J6"].Text = ContactSituationList[i].Info.Year;
            sheet.Range["J8"].Text = ContactSituationList[i].Info.Cooperation_Mode;
            sheet.Range["J10"].Text = ContactSituationList[i].Info.Status;

            List<ContactStatus> contactStatusList = ContactSituationList[i].Status;
            for (int j = 0; j < contactStatusList.Count; j++)
            {
                sheet.Range[j + 13, 2].Text = contactStatusList[j].Contact_Date;
                var mergecolumn = sheet.Merge(sheet.Range[("C" + (13 + j))], sheet.Range["D" + (13 + j)]);
                mergecolumn.HorizontalAlignment = HorizontalAlignType.Left;
                mergecolumn.Text = contactStatusList[j].Contact_Status;
                sheet.Range["E" + (13 + j) + ":K" + (13 + j)].Merge();
                sheet.Range[j + 13, 5].Text = contactStatusList[j].Remarks;
            }

            sheet.Range["B12:K" + (12 + contactStatusList.Count)].BorderAround(LineStyleType.Medium);
            sheet.Range["B12:K" + (12 + contactStatusList.Count)].BorderInside(LineStyleType.Thin);
            return sheet;
        }

        /// <summary>
        /// 如果字串最後一個字為","則將它移除
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string RemoveEndWithComma(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }

            return str.EndsWith(",") ? str.Remove(str.Length - 1) : str;
        }
    }
}
