﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ShareClassLibrary;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using TalentClassLibrary.Model;
using System.Diagnostics;

namespace TalentClassLibrary
{
    /// <summary>
    /// 操作資料庫類別
    /// </summary>
    public class Talent : SQL
    {
        private static Talent talent = new Talent();

        public static Talent GetInstance() => talent;

        /// <summary>
        /// 紀錄大略發生的錯誤
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 刪除指定的面談資料
        /// </summary>
        /// <param name="interviewId">面談ID</param>
        /// <returns></returns>
        public string DelInterviewDataByInterviewId(string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "此面談資料不存在";
            }

            if (this.DelImageByInterviewId(interviewId) == "圖片刪除失敗")
            {

                return "圖片刪除失敗";
            }

            if (TalentFiles.GetInstance().DelFilesByInterviewId(interviewId) == "附加檔案刪除失敗")
            {
                return "附加檔案刪除失敗";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    ////更新最後修改時間
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    ////專案經驗資料
                    cmd.CommandText = @"delete from Project_Experience where Interview_Id = @id";
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    ////面談評語資料
                    cmd.CommandText = @"delete from Interview_Comments where Interview_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////附加檔案
                    cmd.CommandText = @"delete from Files where Interview_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////面談基本資料
                    cmd.CommandText = @"delete from Interview_Info where Interview_Id = @id";
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }

                return "刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "資料庫發生錯誤";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據聯繫ID刪除跟其有關的所有資料
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string DelTalentById(string id)
        {
            if (this.DelImageByContactId(id) == "圖片刪除失敗")
            {
                return "圖片刪除失敗";
            }

            if (this.DelFilesByContactId(id) == "附加檔案刪除失敗")
            {
                return "附加檔案刪除失敗";
            }

            try
            {
                string del = @"delete from Code where Contact_Id = @id"; ////代碼資料
                using (SqlCommand cmd = new SqlCommand(del, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    cmd.ExecuteNonQuery();
                    ////聯繫基本資料
                    cmd.CommandText = @"delete from Contact_Info where Contact_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////聯繫狀況資料
                    cmd.CommandText = @"delete from Contact_Situation where Contact_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////專案經驗資料
                    cmd.CommandText = @"delete from Project_Experience where Interview_Id in (select Interview_Id from Interview_Info where Contact_Id = @id)";
                    cmd.ExecuteNonQuery();
                    ////面談評語資料
                    cmd.CommandText = @"delete from Interview_Comments where Interview_Id in (select Interview_Id from Interview_Info where Contact_Id = @id)";
                    cmd.ExecuteNonQuery();
                    ////附加檔案
                    cmd.CommandText = @"delete from Files where Interview_Id in (select Interview_Id from Interview_Info where Contact_Id = @id)";
                    cmd.ExecuteNonQuery();
                    ////面談基本資料
                    cmd.CommandText = @"delete from Interview_Info where Contact_Id = @id";
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();

                }
                return "刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "資料庫發生錯誤";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除對應聯繫ID的代碼
        /// </summary>
        /// <param name="id">聯繫ID</param>
        private void DelCodeById(string id)
        {
            string del = @"delete from Code where Contact_Id = @id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(del, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 將指定帳號的狀態更改成啟用或停用
        /// </summary>
        /// <param name="account">欲更改狀態的帳號</param>
        /// <returns>啟用(停用)成功(失敗)</returns>
        public string UpdateMemberStatesByAccount(string account, string states)
        {
            string update = @"update Member set States=@states where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.Parameters.Add("@states", SqlDbType.NVarChar).Value = states;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return states + "成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return states + "失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 將資料依狀態分類成修改，新增，刪除的DataTable
        /// </summary>
        /// <param name="DataList">資料</param>
        /// <returns>由新增，修改，刪除DataTable組成的DataSet</returns>
        public DataSet DataDataClassification(System.Data.DataTable DataList)
        {
            DataSet ds = new DataSet();
            DataTable insertList = new DataTable();
            DataTable updateList = new DataTable();
            DataTable delList = new DataTable();

            try
            {
                ////需要新增的資料
                var filter = (from row in DataList.AsEnumerable()
                              where row.Field<string>("Flag") == "I"
                              select row);
                if (filter.Any())
                {
                    insertList = filter.CopyToDataTable();
                }
                ////需要修改的資料
                filter = (from row in DataList.AsEnumerable()
                          where row.Field<string>("Flag") == "U"
                          select row);
                if (filter.Any())
                {
                    updateList = filter.CopyToDataTable();
                }
                ////需要刪除的資料
                filter = (from row in DataList.AsEnumerable()
                          where row.Field<string>("Flag") == "D"
                          select row);
                if (filter.Any())
                {
                    delList = filter.CopyToDataTable();
                }
                ds.Tables.Add(insertList);
                ds.Tables.Add(updateList);
                ds.Tables.Add(delList);
                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return new DataSet();
            }
        }

        /// <summary>
        /// 儲存聯繫狀況資料
        /// </summary>
        /// <param name="data">聯繫狀況資料</param>
        /// <param name="id">對應的聯繫ID</param>
        /// <returns>聯繫資料儲存成功or錯誤訊息</returns>
        public string ContactSituationAction(DataTable data, string id)
        {
            DataSet ds = this.DataDataClassification(data);
            string sqlStr = string.Empty;
            if (!this.ValidIdIsAppear(id))
            {
                return "此聯繫狀況資料，沒有對應的聯繫資料";
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.InsertData(ds.Tables[0], id, "Contact_Situation") != "新增成功")
                {
                    return "聯繫狀況資料新增失敗";
                }
            }

            if (ds.Tables[1].Rows.Count > 0)
            {
                if (this.UpdateData(ds.Tables[1], "Contact_Situation") != "修改成功")
                {
                    return "聯繫狀況資料修改失敗";
                }
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.DelData(ds.Tables[2], "Contact_Situation") != "刪除成功")
                {
                    return "聯繫狀況資料刪除失敗";
                }
            }

            return "聯繫資料儲存成功";
        }

        /// <summary>
        /// 儲存代碼
        /// </summary>
        /// <param name="data">代碼</param>
        /// <param name="id">對應的聯繫ID</param>
        /// <returns>代碼儲存成功or錯誤訊息</returns>
        public string CodeAction(DataTable data, string id)
        {
            DataSet ds = this.DataDataClassification(data);
            string sqlStr = string.Empty;
            if (!this.ValidIdIsAppear(id))
            {
                return "此代碼，沒有對應的聯繫資料";
            }

            string validMsg = this.ValidCodeIsRepeat(data);
            if (!validMsg.Equals(string.Empty))
            {
                return validMsg;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.InsertData(ds.Tables[0], id, "Code") != "新增成功")
                {
                    return "代碼新增失敗";
                }
            }

            if (ds.Tables[1].Rows.Count > 0)
            {
                if (this.UpdateData(ds.Tables[1], "Code") != "修改成功")
                {
                    return "代碼資料修改失敗";
                }
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.DelData(ds.Tables[2], "Code") != "刪除成功")
                {
                    return "代碼刪除失敗";
                }
            }

            return "代碼儲存成功";
        }

        /// <summary>
        /// 根據Table執行刪除動作
        /// </summary>
        /// <param name="DelData">要刪除的資料</param>
        /// <param name="tableName">要刪除的資料</param>
        /// <returns></returns>
        public string DelData(DataTable DelData, string tableName)
        {
            string del = this.CombinationDelSQL(tableName);
            try
            {
                using (SqlCommand cmd = new SqlCommand(del, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in DelData.Rows)
                    {
                        cmd.Parameters.Clear();
                        this.CombinationDelParameters(cmd, dr, tableName);
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "刪除成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據Table組合Parameters(刪除)
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="dr"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private SqlCommand CombinationDelParameters(SqlCommand cmd, DataRow dr, string tableName)
        {
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = dr["Contact_status_Id"].ToString();
                return cmd;
            }
            ////代碼
            if (tableName == "Code")
            {
                cmd.Parameters.Add("@id", SqlDbType.VarChar).Value = dr["Code_Id"].ToString();
                return cmd;
            }

            return cmd;
        }

        /// <summary>
        /// 根據Table產生對應的刪除SQL
        /// </summary>
        /// <param name="tableName">要刪除的Table</param>
        /// <returns>刪除SQL</returns>
        private string CombinationDelSQL(string tableName)
        {
            string Del = string.Empty;
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                Del = @"delete from Contact_Situation where Contact_status_Id = @id";
                return Del;
            }
            ////代碼
            if (tableName == "Code")
            {
                Del = @"delete from Code where Code_Id = @id";
                return Del;
            }

            return Del;
        }

        /// <summary>
        /// 根據Table執行修改動作
        /// </summary>
        /// <param name="updateData">要修改的資料</param>
        /// <returns>修改成功或失敗</returns>
        public string UpdateData(DataTable updateData, string tableName)
        {
            string update = this.CombinationUpdateSQL(tableName);
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in updateData.Rows)
                    {
                        cmd.Parameters.Clear();
                        this.CombinationUpdateParameters(cmd, dr, tableName);
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據Table組合Parameters(修改)
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="dr"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private SqlCommand CombinationUpdateParameters(SqlCommand cmd, DataRow dr, string tableName)
        {
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                cmd.Parameters.Add("@contactStatus", SqlDbType.NVarChar).Value = dr["Contact_Status"].ToString();
                cmd.Parameters.Add("@remarks", SqlDbType.NVarChar).Value = dr["Remarks"].ToString();
                cmd.Parameters.Add("@contactDate", SqlDbType.DateTime).Value = dr["Contact_Date"].ToString();
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = dr["Contact_status_Id"].ToString();
                return cmd;
            }
            ////代碼
            if (tableName == "Code")
            {
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = dr["Id"].ToString();
                cmd.Parameters.Add("@CodeId", SqlDbType.VarChar).Value = dr["Code_Id"].ToString();
                return cmd;
            }

            return cmd;
        }

        /// <summary>
        /// 根據Table產生對應的修改SQL
        /// </summary>
        /// <param name="tableName">要修改的Table</param>
        /// <returns>修改SQL</returns>
        private string CombinationUpdateSQL(string tableName)
        {
            string update = string.Empty;
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                update = @"update Contact_Situation set Contact_Status=@contactStatus,Remarks=@remarks,Contact_Date=@contactDate
                                                       where Contact_status_Id=@id";
                return update;
            }
            ////代碼
            if (tableName == "Code")
            {
                update = @"update Code set Code_Id=@CodeId where Id=@Id";
                return update;
            }

            return update;
        }

        /// <summary>
        /// 根據Table產生對應的新增SQL
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private string CombinationInsertSQL(string tableName)
        {
            string insert = string.Empty;
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                insert = @"insert into Contact_Situation (Contact_Status,Remarks,Contact_Date,Contact_Id)
                                                     values (@contactStatus,@remarks,@contactDate,@contactId)";
                return insert;
            }
            ////代碼
            if (tableName == "Code")
            {
                insert = @"insert into Code (Code_Id,Contact_Id) values (@codeId,@contactId)";
                return insert;
            }

            return insert;
        }

        /// <summary>
        /// 根據Table組合Parameters(新增)
        /// </summary>
        /// <param name="id"></param>
        /// <param name="cmd"></param>
        /// <param name="dr"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private SqlCommand CombinationInsertParameters(string id, SqlCommand cmd, DataRow dr, string tableName)
        {
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                cmd.Parameters.Add("@contactStatus", SqlDbType.NVarChar).Value = dr["Contact_Status"].ToString();
                cmd.Parameters.Add("@remarks", SqlDbType.NVarChar).Value = dr["Remarks"].ToString();
                cmd.Parameters.Add("@contactDate", SqlDbType.DateTime).Value = dr["Contact_Date"].ToString();
                cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = id;
                return cmd;
            }
            ////代碼
            if (tableName == "Code")
            {
                cmd.Parameters.Add("@codeId", SqlDbType.NVarChar).Value = dr["Code_Id"].ToString();
                cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = id;
                return cmd;
            }

            return cmd;
        }

        /// <summary>
        /// 根據Table執行新增動作
        /// </summary>
        /// <param name="inData">要新增的資料</param>
        /// <param name="id">資料對應的聯繫ID或面試ID</param>
        /// <param name="tableName">要新增至哪個Table</param>
        /// <returns>新增成功或失敗</returns>
        public string InsertData(DataTable inData, string id, string tableName)
        {
            string insert = this.CombinationInsertSQL(tableName);
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in inData.Rows)
                    {
                        cmd.Parameters.Clear();
                        this.CombinationInsertParameters(id, cmd, dr, tableName);
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "新增成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 將附加檔案資訊寫入資料庫
        /// </summary>
        /// <param name="files"></param>
        /// <param name="interviewId"></param>
        /// <param name="attachedFileMode"></param>
        /// <returns></returns>
        public string SaveAttachedFilesToDB(List<string> files, string interviewId, string attachedFileMode)
        {
            string sqlStr = string.Empty;
            List<string> path = TalentFiles.GetInstance().SaveAttachedFiles(files, interviewId, attachedFileMode);
            if (path.Contains("上傳失敗"))
            {
                return "檔案上傳失敗";
            }

            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Files where Interview_Id = @Interview_Id and Belong = @Belong";
                    cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                    cmd.Parameters.Add("@Belong", SqlDbType.NVarChar).Value = attachedFileMode;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = @"insert into Files (Interview_Id,File_Path,Belong) values (@Interview_Id,@File_Path,@Belong)";
                    foreach (string file in path)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.Parameters.Add("@File_Path", SqlDbType.NVarChar).Value = file;
                        cmd.Parameters.Add("@Belong", SqlDbType.NVarChar).Value = attachedFileMode;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    this.CommitTransaction();
                    return "檔案上傳成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "檔案上傳失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 將圖片寫入資料庫
        /// </summary>
        /// <param name="interviewId">面談ID</param>
        /// <param name="path">server端路徑</param>
        /// <returns></returns>
        public string UpdateImageByInterviewId(string interviewId, string path)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "此面試基本資料，沒有對應的聯繫資料";
            }

            string update = @"update Interview_Info set Image = @image where Interview_Id = @Interview_Id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@image", SqlDbType.NVarChar).Value = path;
                    cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    this.CommitTransaction();
                }

                return "圖片上傳成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "圖片上傳失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 新增面試基本資料
        /// </summary>
        /// <param name="inData"></param>
        /// <param name="contactId"></param>
        /// <returns></returns>
        public string InsertInterviewInfoData(DataTable inData, string contactId)
        {
            if (!this.ValidIdIsAppear(contactId))
            {
                return "此面試基本資料，沒有對應的聯繫資料";
            }

            string interviewId = string.Empty;
            string insert = @"insert into Interview_Info ([Vacancies],[Interview_Date],[Contact_Id],[Name],[Sex],[Birthday],[Married],[Mail]
                              ,[Adress],[CellPhone],[Expertise_Language],[Expertise_Tools],[Expertise_Tools_Framwork],[Expertise_Devops],[Expertise_OS]
                              ,[Expertise_BigData],[Expertise_DataBase],[Expertise_Certification],[IsStudy],[IsService],[Relatives_Relationship]
                              ,[Relatives_Name],[Care_Work],[Hope_Salary],[When_Report],[Advantage],[Disadvantages],[Hobby],[Attract_Reason]
                              ,[Future_Goal],[Hope_Supervisor],[Hope_Promise],[Introduction],[During_Service]
                              ,[Exemption_Reason],[Urgent_Contact_Person],[Urgent_Relationship],[Urgent_CellPhone],[Education],[Language],[Work_Experience])
                               output inserted.Interview_Id values 
                              (@vacancies,@interviewDate,@contactId,@name,@sex,@birthday,@married,@mail,@adress,@cellPhone,@expertiseLanguage
                              ,@expertiseTools,@expertiseToolsFramwork,@expertiseDevops,@expertiseOS,@expertiseBigData,@expertiseDataBase,@expertiseCertification
                              ,@isStudy,@IsService,@relativesRelationship,@relativesName,@careWork,@hopeSalary,@whenReport,@advantage
                              ,@disadvantages,@hobby,@attractReason,@futureGoal,@hopeSupervisor,@hopePromise,@introduction,@duringService
                              ,@exemptionReason,@urgentContactPerson,@urgentRelationship,@urgentCellPhone,@education,@language,@workExperience)";
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in inData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@education", SqlDbType.NVarChar).Value = dr["Education"].ToString();
                        cmd.Parameters.Add("@language", SqlDbType.NVarChar).Value = dr["Language"].ToString();
                        cmd.Parameters.Add("@workExperience", SqlDbType.NVarChar).Value = dr["Work_Experience"].ToString();
                        cmd.Parameters.Add("@vacancies", SqlDbType.NVarChar).Value = dr["Vacancies"].ToString();
                        cmd.Parameters.Add("@urgentCellPhone", SqlDbType.VarChar).Value = dr["Urgent_CellPhone"].ToString();
                        cmd.Parameters.Add("@urgentRelationship", SqlDbType.NVarChar).Value = dr["Urgent_Relationship"].ToString();
                        cmd.Parameters.Add("@urgentContactPerson", SqlDbType.NVarChar).Value = dr["Urgent_Contact_Person"].ToString();
                        cmd.Parameters.Add("@exemptionReason", SqlDbType.NVarChar).Value = dr["Exemption_Reason"].ToString();
                        cmd.Parameters.Add("@duringService", SqlDbType.NVarChar).Value = dr["During_Service"].ToString();
                        cmd.Parameters.Add("@introduction", SqlDbType.NVarChar).Value = dr["Introduction"].ToString();
                        cmd.Parameters.Add("@hopePromise", SqlDbType.NVarChar).Value = dr["Hope_Promise"].ToString();
                        cmd.Parameters.Add("@hopeSupervisor", SqlDbType.NVarChar).Value = dr["Hope_Supervisor"].ToString();
                        cmd.Parameters.Add("@futureGoal", SqlDbType.NVarChar).Value = dr["Future_Goal"].ToString();
                        cmd.Parameters.Add("@attractReason", SqlDbType.NVarChar).Value = dr["Attract_Reason"].ToString();
                        cmd.Parameters.Add("@hobby", SqlDbType.NVarChar).Value = dr["Hobby"].ToString();
                        cmd.Parameters.Add("@disadvantages", SqlDbType.NVarChar).Value = dr["Disadvantages"].ToString();
                        cmd.Parameters.Add("@advantage", SqlDbType.NVarChar).Value = dr["Advantage"].ToString();
                        cmd.Parameters.Add("@whenReport", SqlDbType.NVarChar).Value = dr["When_Report"].ToString();
                        cmd.Parameters.Add("@hopeSalary", SqlDbType.NVarChar).Value = dr["Hope_salary"].ToString();
                        cmd.Parameters.Add("@careWork", SqlDbType.NVarChar).Value = dr["Care_Work"].ToString();
                        cmd.Parameters.Add("@relativesName", SqlDbType.NVarChar).Value = dr["Relatives_Name"].ToString();
                        cmd.Parameters.Add("@relativesRelationship", SqlDbType.NVarChar).Value = dr["Relatives_Relationship"].ToString();
                        cmd.Parameters.Add("@isService", SqlDbType.NVarChar).Value = dr["IsService"].ToString();
                        cmd.Parameters.Add("@isStudy", SqlDbType.NVarChar).Value = dr["IsStudy"].ToString();
                        cmd.Parameters.Add("@expertiseCertification", SqlDbType.NVarChar).Value = dr["Expertise_Certification"].ToString();
                        cmd.Parameters.Add("@expertiseDataBase", SqlDbType.NVarChar).Value = dr["Expertise_DataBase"].ToString();
                        cmd.Parameters.Add("@expertiseBigData", SqlDbType.NVarChar).Value = dr["Expertise_BigData"].ToString();
                        cmd.Parameters.Add("@expertiseOS", SqlDbType.NVarChar).Value = dr["Expertise_OS"].ToString();
                        cmd.Parameters.Add("@expertiseDevops", SqlDbType.NVarChar).Value = dr["Expertise_Devops"].ToString();
                        cmd.Parameters.Add("@expertiseTools", SqlDbType.NVarChar).Value = dr["Expertise_Tools"].ToString();
                        cmd.Parameters.Add("@expertiseToolsFramwork", SqlDbType.NVarChar).Value = dr["Expertise_Tools_Framwork"].ToString();
                        cmd.Parameters.Add("@expertiseLanguage", SqlDbType.NVarChar).Value = dr["Expertise_language"].ToString();
                        cmd.Parameters.Add("@interviewDate", SqlDbType.Date).Value = dr["Interview_Date"].ToString();
                        cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = contactId;
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString();
                        cmd.Parameters.Add("@birthday", SqlDbType.Date).Value = Common.GetInstance().ValueIsNullOrEmpty(dr["Birthday"].ToString());
                        cmd.Parameters.Add("@married", SqlDbType.NVarChar).Value = dr["Married"].ToString();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString();
                        cmd.Parameters.Add("@adress", SqlDbType.NVarChar).Value = dr["Adress"].ToString();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = dr["CellPhone"].ToString();
                        interviewId = cmd.ExecuteScalar().ToString();
                        if (!string.IsNullOrEmpty(dr["Image"].ToString()))
                        {
                            string path = TalentFiles.GetInstance().UpLoadImage(dr["Image"].ToString(), interviewId);
                            if (path.Equals("上傳失敗") || path.Equals("不存在的路徑"))
                            {
                                throw new Exception(path);
                            }

                            cmd.CommandText = @"update Interview_Info set Image = @image where Interview_Id = @Interview_Id";
                            cmd.Parameters.Add("@image", SqlDbType.NVarChar).Value = path;
                            cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                            cmd.ExecuteNonQuery();
                        }
                    }

                    this.UpdateEditTime(contactId, cmd);
                    this.CommitTransaction();
                }

                return interviewId;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 修改最後編輯時間
        /// </summary>
        /// <param name="contactId"></param>
        /// <param name="cmd"></param>
        private void UpdateEditTime(string contactId, SqlCommand cmd)
        {
            cmd.Parameters.Clear();
            cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = @contactId";
            cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
            cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = contactId;
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 新增聯繫基本資料
        /// </summary>
        /// <param name="inData"></param>
        /// <returns></returns>
        public string InsertContactSituationInfoData(DataTable inData)
        {
            string contact_Id = string.Empty;
            string insert = @"insert into Contact_Info ([Name],[Sex],[Mail],[CellPhone],[UpdateTime],[Cooperation_Mode],[Status],[Place],[Skill],[Year])
                              output inserted.Contact_Id
                              values(@name,@sex,@mail,@cellPhone,@updateTime,@cooperationMode,@status,@place,@skill,@year)";
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in inData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = dr["CellPhone"].ToString();
                        cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        cmd.Parameters.Add("@cooperationMode", SqlDbType.NVarChar).Value = dr["Cooperation_Mode"].ToString();
                        cmd.Parameters.Add("@status", SqlDbType.NVarChar).Value = dr["Status"].ToString();
                        cmd.Parameters.Add("@place", SqlDbType.NVarChar).Value = dr["Place"].ToString();
                        cmd.Parameters.Add("@skill", SqlDbType.NVarChar).Value = dr["Skill"].ToString();
                        cmd.Parameters.Add("@year", SqlDbType.VarChar).Value = dr["Year"].ToString();
                        contact_Id = cmd.ExecuteScalar().ToString();
                    }

                    this.CommitTransaction();
                    return contact_Id;
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 新增使用者帳號
        /// </summary>
        /// <param name="account">欲新增的帳號</param>
        /// <returns>新增成功或失敗</returns>
        public string InsertMember(string account)
        {
            try
            {
                string msg = this.ValidInsertMember(account);
                if (!msg.Equals(string.Empty))
                {
                    return msg;
                }

                account =  TalentCommon.GetInstance().AddMailFormat(account);
                string[] splitAccount = account.Split('@');
                string password = Common.GetInstance().PasswordEncryption(splitAccount[0].ToLower());
                string insert = @"insert into Member (Account,Password,States) values (@account,@password,N'啟用')";
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.Parameters.Add("@password", SqlDbType.VarChar).Value = password;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "新增成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
        }

        /// <summary>
        /// 儲存面談結果資料
        /// </summary>
        /// <param name="saveData">面談結果資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveInterviewResult(DataSet saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"update Interview_Info set Appointment=@Appointment,Results_Remark=@Results_Remark where Interview_Id=@Interview_Id";
                    foreach (DataRow dr in saveData.Tables[1].Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Appointment", SqlDbType.NVarChar).Value = dr["Appointment"].ToString();
                        cmd.Parameters.Add(@"Results_Remark", SqlDbType.NVarChar).Value = dr["Results_Remark"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"delete from [Interview_Comments] where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into [Interview_Comments] ([Interviewer],[Result],[Interview_Id])
                                        values (@Interviewer,@Result,@Interview_Id)";
                    foreach (DataRow dr in saveData.Tables[0].Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Interviewer", SqlDbType.NVarChar).Value = dr["Interviewer"].ToString();
                        cmd.Parameters.Add(@"Result", SqlDbType.NVarChar).Value = dr["Result"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存聯繫狀況資料
        /// </summary>
        /// <param name="saveData">聯繫狀況資料</param>
        /// <param name="contactId">聯繫ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveContactSituation(DataTable saveData, string contactId)
        {
            if (!this.ValidIdIsAppear(contactId))
            {
                return "此聯繫狀況資料，沒有對應的聯繫資料";

            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Contact_Situation where Contact_Id=@Contact_Id";
                    cmd.Parameters.Add("@Contact_Id", SqlDbType.Int).Value = contactId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Contact_Situation (Contact_Status,Remarks,Contact_Date,Contact_Id)
                                        values (@Contact_Status,@Remarks,@Contact_Date,@Contact_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Contact_Status", SqlDbType.NVarChar).Value = dr["Contact_Status"].ToString();
                        cmd.Parameters.Add(@"Remarks", SqlDbType.NVarChar).Value = dr["Remarks"].ToString();
                        cmd.Parameters.Add(@"Contact_Date", SqlDbType.DateTime).Value = dr["Contact_Date"].ToString();
                        cmd.Parameters.Add(@"Contact_Id", SqlDbType.Int).Value = contactId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存代碼資料
        /// </summary>
        /// <param name="saveData">代碼資料</param>
        /// <param name="contactId">聯繫ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveCode(DataTable saveData, string contactId)
        {
            if (!this.ValidIdIsAppear(contactId))
            {
                return "此代碼，沒有對應的聯繫資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Code where Contact_Id=@Contact_Id";
                    cmd.Parameters.Add("@Contact_Id", SqlDbType.Int).Value = contactId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Code (Contact_Id,Code_Id)
                                        values (@Contact_Id,@Code_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Code_Id", SqlDbType.VarChar).Value = dr["Code_Id"].ToString();
                        cmd.Parameters.Add(@"Contact_Id", SqlDbType.Int).Value = contactId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存專案經驗資料
        /// </summary>
        /// <param name="saveData">專案經驗資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveProjectExperience(DataTable saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from [Project_Experience] where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into [Project_Experience] ([Company],[Project_Name],[OS],[Database],[Position],[Language],[Tools],[Description],[Start_End_Date],[Interview_Id])
                                        values (@Company,@Project_Name,@OS,@Database,@Position,@Language,@Tools,@Description,@Start_End_Date,@Interview_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Company", SqlDbType.NVarChar).Value = dr["Company"].ToString();
                        cmd.Parameters.Add(@"Project_Name", SqlDbType.NVarChar).Value = dr["Project_Name"].ToString();
                        cmd.Parameters.Add(@"OS", SqlDbType.NVarChar).Value = dr["OS"].ToString();
                        cmd.Parameters.Add(@"Database", SqlDbType.NVarChar).Value = dr["Database"].ToString();
                        cmd.Parameters.Add(@"Position", SqlDbType.NVarChar).Value = dr["Position"].ToString();
                        cmd.Parameters.Add(@"Language", SqlDbType.NVarChar).Value = dr["Language"].ToString();
                        cmd.Parameters.Add(@"Tools", SqlDbType.NVarChar).Value = dr["Tools"].ToString();
                        cmd.Parameters.Add(@"Description", SqlDbType.NVarChar).Value = dr["Description"].ToString();
                        cmd.Parameters.Add(@"Start_End_Date", SqlDbType.NVarChar).Value = dr["Start_End_Date"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據地點，技能，合作模式，聯繫狀態，最後編輯時間，查詢符合的Id
        /// </summary>
        /// <param name="places">地點，多筆請用","隔開</param>
        /// <param name="expertises">技能，多筆請用","隔開</param>
        /// <param name="cooperationMode">合作模式</param>
        /// <param name="states">聯繫狀態</param>
        /// <param name="startEditDate">起始日期，日期格式</param>
        /// <param name="endEditDate">結束日期，日期格式</param>
        /// <returns>回傳符合條件的Id</returns>
        public List<string> SelectIdByContact(string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            List<string> idList = new List<string>();
            try
            {
                if (Valid.GetInstance().ValidDateRange(startEditDate, endEditDate) != string.Empty)
                {
                    MessageBox.Show("最後編輯日之日期格式或者是日期區間不正確", "錯誤訊息");
                    return new List<string>();
                }

                string[] place = places.Split(',');
                string[] expertise = expertises.Split(',');
                string select = @"select Contact_Id from Contact_Info where ISNULL(Status,'NA') = ISNULL(ISNULL(@status,Status),'NA') and
                                                                            UpdateTime >= ISNULL(@startEditDate, UpdateTime) and
                                                                            UpdateTime <= ISNULL(@endEditDate, UpdateTime) and
                                                                            ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode,Cooperation_Mode),'NA')";
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    ////如果合作模式為"全職"or"合約"，則值為"皆可"也要被查詢出來
                    if (cooperationMode.Equals("全職") || cooperationMode.Equals("合約"))
                    {
                        da.SelectCommand.CommandText += @" or ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode1,Cooperation_Mode),'NA')";
                        da.SelectCommand.Parameters.Add("@CooperationMode1", SqlDbType.NChar).Value = Common.GetInstance().ValueIsNullOrEmpty("皆可");
                    }
                    ////多筆地點
                    for (int i = 0; i < place.Length; i++)
                    {
                        if (i == 0)
                        {
                            da.SelectCommand.CommandText += @" and ISNULL(Place,'NA') LIKE ISNULL(ISNULL(@place" + (i + 1) + ", Place),'NA')";
                            da.SelectCommand.Parameters.Add("@place" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + place[i] + "%");
                        }
                        else
                        {
                            da.SelectCommand.CommandText += @" or ISNULL(Place,'NA') LIKE ISNULL(ISNULL(@place" + (i + 1) + ", Place),'NA')";
                            da.SelectCommand.Parameters.Add("@place" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + place[i] + "%");
                        }
                    }
                    ////多筆技能
                    for (int i = 0; i < expertise.Length; i++)
                    {
                        if (i == 0)
                        {
                            da.SelectCommand.CommandText += @" and ISNULL(Skill,'NA') Like ISNULL(ISNULL(@skill" + (i + 1) + ", Skill),'NA')";
                            da.SelectCommand.Parameters.Add("@skill" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + expertise[i] + "%");
                        }
                        else
                        {
                            da.SelectCommand.CommandText += @" or ISNULL(Skill,'NA') Like ISNULL(ISNULL(@skill" + (i + 1) + ", Skill),'NA')";
                            da.SelectCommand.Parameters.Add("@skill" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + expertise[i] + "%");
                        }
                    }

                    da.SelectCommand.Parameters.Add("@CooperationMode", SqlDbType.NChar).Value = TalentCommon.GetInstance().ValueIsAny(cooperationMode);
                    da.SelectCommand.Parameters.Add("@status", SqlDbType.NChar).Value = TalentCommon.GetInstance().ValueIsAny(states);
                    da.SelectCommand.Parameters.Add("@startEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startEditDate);
                    da.SelectCommand.Parameters.Add("@endEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endEditDate);


                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            idList.Add(dr[0].ToString());
                        }
                    }

                    return idList;
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new List<string>();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        ///根據關鍵字查詢符合資料的ID
        /// </summary>
        /// <param name="keyWords">關鍵字，如有多個關鍵字請用","隔開</param>
        /// <returns>符合關鍵字的ID</returns>
        public List<string> SelectIdByKeyWord(string keyWords)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            List<string> idList = new List<string>();
            try
            {
                string select = string.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    ////關鍵字為空則撈出所有ID
                    if (string.IsNullOrEmpty(keyWords))
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Contact_Info)";
                        da.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                idList.Add(dr[0].ToString());
                            }
                        }

                        return idList;
                    }

                    string[] keyWord = keyWords.Split(',');
                    foreach (string word in keyWord)
                    {
                        if (string.IsNullOrEmpty(word))
                        {
                            continue;
                        }

                        da.SelectCommand.Parameters.Clear();
                        da.SelectCommand.CommandType = CommandType.StoredProcedure;
                        da.SelectCommand.CommandText = "KeyWordSearch";
                        da.SelectCommand.Parameters.Add("@KeyWord", SqlDbType.NVarChar).Value = word;
                        Common.GetInstance().ClearDataTable(dt);
                        da.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            idList.AddRange(this.SelectIdbyResult(dt, da));
                        }
                    }

                    return idList.Distinct().ToList();
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new List<string>();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據面談條件查詢符合的聯繫ID
        /// </summary>
        /// <param name="isInterview">是否已面談，值為"已面談","未面談","不限"</param>
        /// <param name="interviewResult">面談結果，值為"錄用","不錄用","暫保留","不限"</param>
        /// <param name="startInterviewDate">起始日期，日期格式</param>
        /// <param name="endInterviewDate">結束日期，日期格式</param>
        /// <returns>回傳符合條件的聯繫ID</returns>
        public List<string> SelectIdByInterviewFilter(string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            List<string> idList = new List<string>();
            string select = string.Empty;
            try
            {
                if (Valid.GetInstance().ValidDateRange(startInterviewDate, endInterviewDate) != string.Empty)
                {
                    MessageBox.Show("面試日期之日期格式或者是日期區間不正確", "錯誤訊息");
                    return new List<string>();
                }

                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    ////是否已面談
                    if (isInterview == "已面談")
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Interview_Info where Appointment is not null and Appointment !='') INTERSECT ";
                    }
                    else if (isInterview == "未面談")
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Contact_Info
                                                          EXCEPT
                                                          select Contact_Id from Interview_Info where Appointment is not null and Appointment !='') INTERSECT ";
                    }
                    ////面談結果
                    if (interviewResult != "不限")
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Interview_Info where Appointment = ISNULL(@interviewResult, Appointment)) INTERSECT ";
                        da.SelectCommand.Parameters.Add("@interviewResult", SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty(interviewResult);
                    }
                    ////面談日期
                    if (!string.IsNullOrEmpty(startInterviewDate) || !string.IsNullOrEmpty(endInterviewDate))
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Interview_Info where Interview_Date <= ISNULL(@endInterviewDate, Interview_Date) AND Interview_Date >= ISNULL(@startInterviewDate, Interview_Date)) INTERSECT ";
                        da.SelectCommand.Parameters.Add("@endInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endInterviewDate);
                        da.SelectCommand.Parameters.Add("@startInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startInterviewDate);
                    }
                    ////所有ID
                    da.SelectCommand.CommandText += @"(select Contact_Id from Contact_Info)";
                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            idList.Add(dr[0].ToString());
                        }
                    }

                    return idList;
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new List<string>();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據條件查詢符合的ID
        /// </summary>
        /// <param name="keyWords">關鍵字，如有多個關鍵字請用","隔開</param>
        /// <param name="places">地點，多筆請用","隔開</param>
        /// <param name="expertises">技能，多筆請用","隔開</param>
        /// <param name="cooperationMode">合作模式</param>
        /// <param name="states">聯繫狀態</param>
        /// <param name="startEditDate">起始日期，日期格式</param>
        /// <param name="endEditDate">結束日期，日期格式</param>
        /// <param name="isInterview">是否已面談，值為"已面談","未面談","不限"</param>
        /// <param name="interviewResult">面談結果，值為"錄用","不錄用","暫保留","不限"</param>
        /// <param name="startInterviewDate">起始日期，日期格式</param>
        /// <param name="endInterviewDate">結束日期，日期格式</param>
        /// <returns></returns>
        public List<string> SelectIdByFilter(string keyWords, string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate, string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate)
        {
            List<string> idListBykeywords = this.SelectIdByKeyWord(keyWords);
            List<string> idListByContact = this.SelectIdByContact(places, expertises, cooperationMode, states, startEditDate, endEditDate);
            List<string> idListByInterviewFilter = this.SelectIdByInterviewFilter(isInterview, interviewResult, startInterviewDate, endInterviewDate);
            List<string> idList = new List<string>();
            idList = idListBykeywords.Intersect(idListByContact).ToList();
            idList = idList.Intersect(idListByInterviewFilter).ToList();
            return idList;

        }

        ///透過預存程序會找出符合值得所在Table and 欄位
        /// *****************************************
        /// *ColumnName               | ColumnValue * 
        /// **************************|**************
        /// *[dbo].[Education].[School]|屏東科技大學*
        /// *****************************************

        /// <summary>
        ///透過Table and 欄位查詢對應的ID
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="da"></param>
        /// <returns></returns>
        private List<string> SelectIdbyResult(DataTable dt, SqlDataAdapter da)
        {
            List<string> idList = new List<string>();
            da.SelectCommand.Parameters.Clear();
            da.SelectCommand.CommandText = string.Empty;
            da.SelectCommand.CommandType = CommandType.Text;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                string[] tableName = dr[0].ToString().Split('.');
                if (tableName[1] == "[Contact_Status_List]" || tableName[1] == "[Member]")
                {
                    continue;
                }

                if (tableName[1] == "[Code]" || tableName[1] == "[Contact_Situation]" || tableName[1] == "[Contact_Info]" || tableName[1] == "[Interview_Info]")
                {
                    if (i == dt.Rows.Count - 1)
                    {
                        da.SelectCommand.CommandText += @"select Contact_Id from " + tableName[1] + " where " + tableName[2] + " = @value" + i;
                    }
                    else
                    {
                        da.SelectCommand.CommandText += @"select Contact_Id from " + tableName[1] + " where " + tableName[2] + " = @value" + i + " UNION ";
                    }

                    da.SelectCommand.Parameters.Add("@value" + i, SqlDbType.NVarChar).Value = dr[1].ToString();
                }
                else
                {
                    if (i == dt.Rows.Count - 1)
                    {
                        da.SelectCommand.CommandText += @"select b.Contact_Id from " + tableName[1] + " a,Interview_Info b where a.Interview_Id = b.Interview_Id and a." + tableName[2] + " = @value" + i;
                    }
                    else
                    {
                        da.SelectCommand.CommandText += @"select b.Contact_Id from " + tableName[1] + " a,Interview_Info b where a.Interview_Id = b.Interview_Id and a." + tableName[2] + " = @value" + i + " UNION ";
                    }

                    da.SelectCommand.Parameters.Add("@value" + i, SqlDbType.NVarChar).Value = dr[1].ToString();
                }
            }

            Common.GetInstance().ClearDataTable(dt);
            if (da.SelectCommand.CommandText.EndsWith(" UNION "))
            {
                da.SelectCommand.CommandText = da.SelectCommand.CommandText.Remove(da.SelectCommand.CommandText.Length - 7);
            }

            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    idList.Add(dr[0].ToString());
                }
            }

            return idList;
        }

        /// <summary>
        /// 根據ID撈出人才資料
        /// </summary>
        /// <param name="idList"></param>
        /// <returns></returns>
        public DataTable SelectTalentInfoById(List<string> idList)
        {
            ErrorMessage = string.Empty;
            string select = string.Empty;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Name");
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Date");
            dt.Columns.Add("Contact_Status");
            dt.Columns.Add("Remarks");
            dt.Columns.Add("Interview_Date");
            dt.Columns.Add("UpdateTime");
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    foreach (string id in idList)
                    {
                        ds.Tables.Clear();
                        da.SelectCommand.Parameters.Clear();
                        da.SelectCommand.CommandText = @"select top 1 Contact_Id,Name,CONVERT(varchar(100), UpdateTime, 111) UpdateTime from Contact_Info where Contact_Id = @Contact_Id ";
                        da.SelectCommand.CommandText += @"select Code_Id from Code where Contact_Id=@Contact_Id ";
                        da.SelectCommand.CommandText += @"select top 2 CONVERT(varchar(100), Contact_Date, 111) Contact_Date,Contact_Status,Remarks from Contact_Situation where Contact_Id = @Contact_Id order by Contact_Date desc ";
                        da.SelectCommand.CommandText += @"select top 2 CONVERT(varchar(100), Interview_Date, 111) Interview_Date from Interview_Info where Contact_Id = @Contact_Id order by Interview_Date desc";
                        da.SelectCommand.Parameters.Add("@Contact_Id", SqlDbType.Int).Value = id;
                        da.Fill(ds);
                        CombinationGrid(ds, dt);
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 將查詢結果組合成符合格式的DataTable
        /// </summary>
        /// <param name="ds">查詢決果</param>
        /// <param name="dt">輸出DataTable</param>
        public static void CombinationGrid(DataSet ds, DataTable dt)
        {
            DataRow row = dt.NewRow();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                row["Contact_Id"] = dr["Contact_Id"].ToString();
                row["Name"] = dr["Name"].ToString();
                row["UpdateTime"] = dr["UpdateTime"].ToString();
            }

            foreach (DataRow dr in ds.Tables[1].Rows)
            {
                row["Code_Id"] += dr[0].ToString() + "\n";
            }

            foreach (DataRow dr in ds.Tables[2].Rows)
            {
                row["Contact_Date"] = dr["Contact_Date"].ToString();
                row["Contact_Status"] = dr["Contact_Status"].ToString();
                row["Remarks"] = dr["Remarks"].ToString();
                break;
            }

            foreach (DataRow dr in ds.Tables[3].Rows)
            {
                row["Interview_Date"] += dr[0].ToString() + "\n";
            }

            dt.Rows.Add(row);

            DataRow row1 = dt.NewRow();

            if (ds.Tables[2].Rows.Count > 1)
            {
                foreach (DataRow dr in ds.Tables[2].Rows)
                {
                    row1["Contact_Date"] = dr["Contact_Date"].ToString();
                    row1["Contact_Status"] = dr["Contact_Status"].ToString();
                    row1["Remarks"] = dr["Remarks"].ToString();
                }

                dt.Rows.Add(row1);
            }
        }

        /// <summary>
        /// 根據聯繫ID查詢相關的聯繫基本資料與聯繫狀況
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public DataSet SelectContactSituationDataById(string id)
        {
            ErrorMessage = string.Empty;
            DataSet ds = new DataSet();
            string select = @"select [Name],[Sex],[Mail],[CellPhone],[Cooperation_Mode],[Status],[Place],[Skill],[Year] from Contact_Info a where a.Contact_Id = @id
                              select CONVERT(varchar(100), a.Contact_Date, 111) Contact_Date,a.Contact_Status,a.Remarks from Contact_Situation a where a.Contact_Id = @id order by Contact_Date desc
                              select a.Code_Id from Code a where a.Contact_Id = @id
                              select CONVERT(varchar(100), UpdateTime, 20) UpdateTime from Contact_Info where Contact_Id = @id
                              select Interview_Id from Interview_Info where Contact_Id = @id";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    da.Fill(ds);
                }

                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataSet();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        public DataTable SelectContactSituationInfoByCode(List<string> CodeList)
        {
            DataTable dt = new DataTable();
            string whereCode = string.Empty;
            try
            {
                for (int i = 0; i < CodeList.Count(); i++)
                {
                    whereCode += "@Code" + i + ", ";
                }

                whereCode = whereCode.Substring(0, whereCode.Length - 2);
                string select = @"select a.Name,a.Code,b.Contact_Date,b.Contact_Status,b.Remarks,a.UpdateTime 
                              from Talent_Info a left join Contact_Situation b on a.Code = b.Code 
                              and a.Code in (select code from Talent_Info where code in (" + whereCode + ") and Visible = 'true')";
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    for (int i = 0; i < CodeList.Count(); i++)
                    {
                        da.SelectCommand.Parameters.Add("@Code" + i, SqlDbType.Int).Value = CodeList[i];
                    }

                    da.Fill(dt);
                    return dt;
                }
            }
            catch
            {
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據面談ID與附加檔案模式來查詢符合的資料
        /// </summary>
        /// <param name="interviewId">面談ID</param>
        /// <param name="filesMode">附加檔案模式 EX：面談基本資料，專案經驗</param>
        /// <returns></returns>
        public DataTable SelectFiles(string interviewId, string filesMode)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select File_Path,Belong from Files where Interview_Id = @Interview_Id and Belong = @Belong";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                    da.SelectCommand.Parameters.Add("@Belong", SqlDbType.NVarChar).Value = filesMode;
                    da.Fill(dt);
                }

                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        ///根據面談資料ID查詢面談資料
        /// </summary>
        /// <param name="interviewId"></param>
        /// <returns></returns>
        public DataSet SelectInterviewDataById(string interviewId)
        {
            ErrorMessage = string.Empty;
            DataSet ds = new DataSet();
            string SQLStr = @"select [Vacancies],CONVERT(varchar(100), Interview_Date, 111)Interview_Date,[Name],[Sex],CONVERT(varchar(100), [Birthday], 111) Birthday,[Married],[Mail],[Adress],[CellPhone],[Image]
                                    ,[Expertise_Language],[Expertise_Tools],[Expertise_Tools_Framwork],[Expertise_Devops],[Expertise_OS],[Expertise_BigData],[Expertise_DataBase]
                                    ,[Expertise_Certification],[IsStudy],[IsService],[Relatives_Relationship],[Relatives_Name],[Care_Work],[Hope_Salary]
                                    ,[When_Report],[Advantage],[Disadvantages],[Hobby],[Attract_Reason],[Future_Goal],[Hope_Supervisor]
                                    ,[Hope_Promise],[Introduction],[During_Service],[Exemption_Reason],[Urgent_Contact_Person],[Urgent_Relationship]
                                    ,[Urgent_CellPhone],[Education],[Language],[Work_Experience] from Interview_Info where Interview_Id=@interviewId
                              select Interviewer,Result from Interview_Comments  where Interview_Id=@interviewId
                              select Appointment,Results_Remark from Interview_Info where Interview_Id=@interviewId
                              select [Company],[Project_Name],[OS],[Database],[Position],[Language],[Tools],[Description],[Start_End_Date] from Project_Experience  where Interview_Id=@interviewId";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(SQLStr, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    da.Fill(ds);
                }
                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataSet();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 搜尋使用者帳號資訊，hr帳號要排除
        /// </summary>
        /// <returns>使用者帳號資訊</returns>
        public DataTable SelectMemberInfo()
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Account,States from Member where account != 'hr@is-land.com.tw'";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.Fill(dt);
                    if (dt.Rows.Count == 0)
                    {
                        return new DataTable();
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據帳號搜尋使用者資訊
        /// </summary>
        /// <param name="account">帳號</param>
        /// <returns>帳號資訊EX：hr@is-land.com.tw,啟用</returns>
        public DataTable SelectMemberInfoByAccount(string account)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Account,States from Member where Account=@account";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@account", SqlDbType.NVarChar).Value = account;
                    da.Fill(dt);
                    if (dt.Rows.Count == 0)
                    {
                        return new DataTable();
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 登入功能
        /// </summary>
        /// <param name="account">帳號</param>
        /// <param name="password">密碼</param>
        /// <returns>登入成功回傳帳號，登入失敗回傳"登入失敗"，該帳號停用中回傳"該帳號停用中"</returns>
        public string SignIn(string account, string password)
        {
            ErrorMessage = string.Empty;
            string msg = "登入失敗";
            account = TalentCommon.GetInstance().AddMailFormat(account);
            password = Common.GetInstance().PasswordEncryption(password.ToLower());
            string select = @"select Account,States from Member where Account=@account and Password = @password";
            DataTable dt = new DataTable();
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@account", SqlDbType.NVarChar).Value = account;
                    da.SelectCommand.Parameters.Add("@password", SqlDbType.NVarChar).Value = password;

                    da.Fill(dt);

                    if (dt.Rows.Count == 0)
                    {
                        return msg;
                    }
                }
                if (TalentValid.GetInstance().ValidMemberStates(dt.Rows[0][1].ToString()))
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return "該帳號停用中";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return msg;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 修改聯繫基本資料
        /// </summary>
        /// <param name="editData">聯繫基本資料</param>
        /// <param name="id">聯繫ID</param>
        /// <returns>修改成功或失敗</returns>
        public string UpdateContactSituationInfoData(DataTable editData, string id)
        {
            if (!this.ValidIdIsAppear(id))
            {
                return "修改失敗";
            }

            string update = @"update Contact_Info set Name=@name,Sex=@sex,Mail=@mail,CellPhone=@cellPhone,UpdateTime=@updateTime,
                                                      Cooperation_Mode=@cooperationMode,Status=@status,Place=@place,Skill=@skill,Year=@year
                              where Contact_Id=@id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in editData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = dr["CellPhone"].ToString();
                        cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        cmd.Parameters.Add("@cooperationMode", SqlDbType.NChar).Value = dr["Cooperation_Mode"].ToString();
                        cmd.Parameters.Add("@status", SqlDbType.NVarChar).Value = dr["Status"].ToString();
                        cmd.Parameters.Add("@place", SqlDbType.NVarChar).Value = dr["Place"].ToString();
                        cmd.Parameters.Add("@skill", SqlDbType.NVarChar).Value = dr["Skill"].ToString();
                        cmd.Parameters.Add("@year", SqlDbType.VarChar).Value = dr["Year"].ToString();
                        cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存學歷資料
        /// </summary>
        /// <param name="saveData">學歷資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveEducation(DataTable saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Education where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Education ([School],[Department],[Start_End_Date],[Is_Graduation],[Remark],[Interview_Id])
                                        values (@School,@Department,@Start_End_Date,@Is_Graduation,@Remark,@Interview_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"School", SqlDbType.NVarChar).Value = dr["School"].ToString();
                        cmd.Parameters.Add(@"Department", SqlDbType.NVarChar).Value = dr["Department"].ToString();
                        cmd.Parameters.Add(@"Start_End_Date", SqlDbType.NVarChar).Value = dr["Start_End_Date"].ToString();
                        cmd.Parameters.Add(@"Is_Graduation", SqlDbType.NChar).Value = dr["Is_Graduation"].ToString();
                        cmd.Parameters.Add(@"Remark", SqlDbType.NVarChar).Value = dr["Remark"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存工作經歷資料
        /// </summary>
        /// <param name="saveData">工作經歷資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveWorkExperience(DataTable saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Work_Experience where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Work_Experience ([Institution_name],[Position],[Start_End_Date],[Start_Salary],[End_Salary],[Leaving_Reason],[Interview_Id])
                                        values (@Institution_name,@Position,@Start_End_Date,@Start_Salary,@End_Salary,@Leaving_Reason,@Interview_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Institution_name", SqlDbType.NVarChar).Value = dr["Institution_name"].ToString();
                        cmd.Parameters.Add(@"Position", SqlDbType.NVarChar).Value = dr["Position"].ToString();
                        cmd.Parameters.Add(@"Start_End_Date", SqlDbType.NVarChar).Value = dr["Start_End_Date"].ToString();
                        cmd.Parameters.Add(@"Start_Salary", SqlDbType.NVarChar).Value = dr["Start_Salary"].ToString();
                        cmd.Parameters.Add(@"End_Salary", SqlDbType.NVarChar).Value = dr["End_Salary"].ToString();
                        cmd.Parameters.Add(@"Leaving_Reason", SqlDbType.NVarChar).Value = dr["Leaving_Reason"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據聯繫ID刪除附加檔案
        /// </summary>
        /// <param name="contactId"></param>
        /// <returns></returns>
        public string DelFilesByContactId(string contactId)
        {
            DataTable dt = new DataTable();
            string select = @"select Interview_Id from Interview_Info where Contact_Id = @id";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = contactId;
                    da.Fill(dt);
                }

                foreach (DataRow dr in dt.Rows)
                {
                    string serverDir = @".\files\" + dr[0].ToString().Trim();
                    if (Directory.Exists(serverDir))
                    {
                        Directory.Delete(serverDir, true);
                    }
                }

                return "附加檔案刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "附加檔案刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除圖片檔 聯繫ID
        /// </summary>
        /// <param name="contactId"></param>
        /// <returns></returns>
        public string DelImageByContactId(string contactId)
        {
            string select = @" select Image from Interview_Info where Contact_Id = @id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = contactId;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            if (File.Exists(dr[0].ToString()))
                            {
                                File.Delete(dr[0].ToString());
                            }
                        }
                    }
                }

                return "圖片刪除成功";
            }
            catch (IOException ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "圖片刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除圖片檔 by面談ID
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string DelImageByInterviewId(string interviewId)
        {
            string select = @"select Image from Interview_Info where Interview_Id = @Interview_Id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            if (File.Exists(dr[0].ToString()))
                            {
                                File.Delete(dr[0].ToString());
                            }
                        }
                    }
                }

                return "圖片刪除成功";
            }
            catch (IOException ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "圖片刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        

        /// <summary>
        /// 修改面試基本資料
        /// </summary>
        /// <param name="editData">面試基本資料</param>
        /// <param name="interviewId">面試ID</param>
        /// <returns></returns>
        public string UpdateInterviewInfoData(DataTable editData, string interviewId, string dbPath, string uiPath)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "修改失敗";
            }

            string path = TalentFiles.GetInstance().UpdateImage(dbPath, uiPath, interviewId);
            if (path.Equals("上傳失敗") || path.Equals("不存在的路徑") || path.Equals("檔案刪除失敗"))
            {
                throw new Exception();
            }

            string update = @"update Interview_Info set Vacancies=@vacancies,Interview_Date=@interviewDate,Name=@name,Sex=@sex,Birthday=@birthday
                              ,Married=@married,Mail=@mail,Adress=@adress,CellPhone=@cellPhone,Image=@image,Expertise_Language=@expertiseLanguage
                              ,Expertise_Tools=@expertiseTools,Expertise_Tools_Framwork=@expertiseToolsFramwork,Expertise_Devops=@expertiseDevops
                              ,Expertise_OS=@expertiseOS
                              ,Expertise_BigData=@expertiseBigData,Expertise_DataBase=@expertiseDataBase,Expertise_Certification=@expertiseCertification
                              ,IsStudy=@isStudy,IsService=@IsService,Relatives_Relationship=@relativesRelationship,Relatives_Name=@relativesName
                              ,Care_Work=@careWork,Hope_Salary=@hopeSalary,When_Report=@whenReport,Advantage=@advantage,Disadvantages=@disadvantages
                              ,Hobby=@hobby,Attract_Reason=@attractReason,Future_Goal=@futureGoal,Hope_Supervisor=@hopeSupervisor
                              ,Hope_Promise=@hopePromise,Introduction=@introduction,During_Service=@duringService,Exemption_Reason=@exemptionReason
                              ,Urgent_Contact_Person=@urgentContactPerson,Urgent_Relationship=@urgentRelationship,Urgent_CellPhone=@urgentCellPhone
                              ,Education=@education,Language=@language,Work_Experience=@workExperience
                              where Interview_Id=@interviewId";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in editData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@education", SqlDbType.NVarChar).Value = dr["Education"].ToString().Trim();
                        cmd.Parameters.Add("@language", SqlDbType.NVarChar).Value = dr["Language"].ToString().Trim();
                        cmd.Parameters.Add("@workExperience", SqlDbType.NVarChar).Value = dr["Work_Experience"].ToString().Trim();
                        cmd.Parameters.Add("@vacancies", SqlDbType.NVarChar).Value = dr["Vacancies"].ToString().Trim();
                        cmd.Parameters.Add("@urgentCellPhone", SqlDbType.Char).Value = dr["Urgent_CellPhone"].ToString().Trim();
                        cmd.Parameters.Add("@urgentRelationship", SqlDbType.NVarChar).Value = dr["Urgent_Relationship"].ToString().Trim();
                        cmd.Parameters.Add("@urgentContactPerson", SqlDbType.NVarChar).Value = dr["Urgent_Contact_Person"].ToString().Trim();
                        cmd.Parameters.Add("@exemptionReason", SqlDbType.NVarChar).Value = dr["Exemption_Reason"].ToString().Trim();
                        cmd.Parameters.Add("@duringService", SqlDbType.NVarChar).Value = dr["During_Service"].ToString().Trim();
                        cmd.Parameters.Add("@introduction", SqlDbType.NVarChar).Value = dr["Introduction"].ToString().Trim();
                        cmd.Parameters.Add("@hopePromise", SqlDbType.NVarChar).Value = dr["Hope_Promise"].ToString().Trim();
                        cmd.Parameters.Add("@hopeSupervisor", SqlDbType.NVarChar).Value = dr["Hope_Supervisor"].ToString().Trim();
                        cmd.Parameters.Add("@futureGoal", SqlDbType.NVarChar).Value = dr["Future_Goal"].ToString().Trim();
                        cmd.Parameters.Add("@attractReason", SqlDbType.NVarChar).Value = dr["Attract_Reason"].ToString().Trim();
                        cmd.Parameters.Add("@hobby", SqlDbType.NVarChar).Value = dr["Hobby"].ToString().Trim();
                        cmd.Parameters.Add("@disadvantages", SqlDbType.NVarChar).Value = dr["Disadvantages"].ToString().Trim();
                        cmd.Parameters.Add("@advantage", SqlDbType.NVarChar).Value = dr["Advantage"].ToString().Trim();
                        cmd.Parameters.Add("@whenReport", SqlDbType.NVarChar).Value = dr["When_Report"].ToString().Trim();
                        cmd.Parameters.Add("@hopeSalary", SqlDbType.NVarChar).Value = dr["Hope_salary"].ToString().Trim();
                        cmd.Parameters.Add("@careWork", SqlDbType.NVarChar).Value = dr["Care_Work"].ToString().Trim();
                        cmd.Parameters.Add("@relativesName", SqlDbType.NVarChar).Value = dr["Relatives_Name"].ToString().Trim();
                        cmd.Parameters.Add("@relativesRelationship", SqlDbType.NVarChar).Value = dr["Relatives_Relationship"].ToString().Trim();
                        cmd.Parameters.Add("@isService", SqlDbType.NVarChar).Value = dr["IsService"].ToString().Trim();
                        cmd.Parameters.Add("@isStudy", SqlDbType.NVarChar).Value = dr["IsStudy"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseCertification", SqlDbType.NVarChar).Value = dr["Expertise_Certification"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseDataBase", SqlDbType.NVarChar).Value = dr["Expertise_DataBase"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseBigData", SqlDbType.NVarChar).Value = dr["Expertise_BigData"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseOS", SqlDbType.NVarChar).Value = dr["Expertise_OS"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseDevops", SqlDbType.NVarChar).Value = dr["Expertise_Devops"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseTools", SqlDbType.NVarChar).Value = dr["Expertise_Tools"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseToolsFramwork", SqlDbType.NVarChar).Value = dr["Expertise_Tools_Framwork"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseLanguage", SqlDbType.NVarChar).Value = dr["Expertise_language"].ToString().Trim();
                        cmd.Parameters.Add("@image", SqlDbType.NVarChar).Value = path.Trim();
                        cmd.Parameters.Add("@interviewDate", SqlDbType.Date).Value = dr["Interview_Date"].ToString().Trim();
                        cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString().Trim();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString().Trim();
                        cmd.Parameters.Add("@birthday", SqlDbType.Date).Value = Common.GetInstance().ValueIsNullOrEmpty(dr["Birthday"].ToString().Trim());
                        cmd.Parameters.Add("@married", SqlDbType.NVarChar).Value = dr["Married"].ToString().Trim();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString().Trim();
                        cmd.Parameters.Add("@adress", SqlDbType.NVarChar).Value = dr["Adress"].ToString().Trim();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.NVarChar).Value = dr["CellPhone"].ToString().Trim();
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        public string UpdatePasswordByaccount(string account, string newPassword)
        {
            string password = Common.GetInstance().PasswordEncryption(newPassword.ToLower());
            string update = @"update Member set Password=@password where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@password", SqlDbType.VarChar).Value = password;
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 更換密碼功能
        /// </summary>
        /// <param name="account">帳號</param>
        /// <param name="oldPassword">舊密碼</param>
        /// <param name="newPassword">新密碼</param>
        /// <param name="checkNewPassword">確認新密碼</param>
        /// <returns>是否修改成功</returns>
        public string UpdatePasswordByaccount(string account, string oldPassword, string newPassword, string checkNewPassword)
        {
            if (!this.SignIn(account, oldPassword).Equals(account))
            {
                return "密碼錯誤";
            }

            string msg = TalentValid.GetInstance().ValidNewPassword(newPassword, checkNewPassword);
            if (!msg.Equals(string.Empty))
            {
                return msg;
            }

            string password = Common.GetInstance().PasswordEncryption(newPassword.ToLower());
            string update = @"update Member set Password=@password where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@password", SqlDbType.VarChar).Value = password;
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 檢查面談資料ID是否存在
        /// </summary>
        /// <param name="InterviewId"></param>
        /// <returns></returns>
        public bool ValidInterviewIdIsAppear(string InterviewId)
        {
            bool msg = false;
            string select = @"select count(1) from Interview_Info where Interview_Id = @interviewId";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = InterviewId;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            if (int.Parse(dr[0].ToString()) > 0)
                            {
                                msg = true;
                            }
                            else
                            {
                                msg = false;
                            }
                        }
                    }
                }
                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return false;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證此聯繫ID是否存在
        /// </summary>
        /// <param name="id">欲驗證的聯繫ID</param>
        /// <returns>True or False</returns>
        public bool ValidIdIsAppear(string id)
        {
            bool msg = false;
            string select = @"select count(1) from Contact_Info where Contact_Id = @contactId";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = id;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            if (int.Parse(dr[0].ToString()) > 0)
                            {
                                msg = true;
                            }
                            else
                            {
                                msg = false;
                            }
                        }
                    }
                }
                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return false;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證從Excel匯入的代碼是否已存在於資料庫
        /// </summary>
        /// <param name="codeList"></param>
        /// <returns></returns>
        public string ValidCodeIsRepeat(List<string> codeList)
        {
            string msg = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Code_Id from Code where Code_Id in (";

            if (codeList.Count == 0)
            {
                return msg;
            }

            for (int i = 0; i < codeList.Count; i++)
            {
                select += @"@codeId" + i + ",";
            }

            select = select.RemoveEndWithDelimiter(",") + ")";

            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    for (int i = 0; i < codeList.Count; i++)
                    {
                        da.SelectCommand.Parameters.Add("@codeId" + i, SqlDbType.VarChar).Value = codeList[i];
                    }

                    da.Fill(dt);
                }

                if(dt.Rows.Count == 0)
                {
                    return msg;
                }

                foreach(DataRow dr in dt.Rows)
                {
                    msg += dr[0].ToString() + "此代碼已存在\n";
                }

                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                msg = "資料庫錯誤";
                return msg;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證欲新增，修改的代碼是否已存在
        /// </summary>
        /// <param name="codeList">代碼陣列</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidCodeIsRepeat(DataTable codeList)
        {
            string msg = string.Empty;
            string select = @"select count(1) from Code where Code_Id = @codeId and Contact_Id != @Contact_Id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    foreach (DataRow row in codeList.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@codeId", SqlDbType.VarChar).Value = row["Code_Id"].ToString();
                        cmd.Parameters.Add("@Contact_Id", SqlDbType.VarChar).Value = row["Contact_Id"].ToString();
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                if (int.Parse(dr[0].ToString()) > 0)
                                {
                                    msg += row["Code_Id"].ToString() + "此代碼已存在\n";
                                }
                            }
                        }
                    }

                }
                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                msg = "發生未預期錯誤";
                return msg;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證欲新增的帳號是否已存在
        /// </summary>
        /// <param name="account">欲新增的帳號</param>
        /// <returns>空值代表不存在</returns>
        public string ValidInsertMember(string account)
        {
            if (string.IsNullOrEmpty(account))
            {
                return "帳號不可為空值";
            }

            int count = 0;
            account = TalentCommon.GetInstance().AddMailFormat(account);
            string msg = TalentValid.GetInstance().ValidIsCompanyMail(account);
            if (!msg.Equals(string.Empty))
            {
                return msg;
            }

            string select = @"select count(1) from Member where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    using (SqlDataReader re = cmd.ExecuteReader())
                    {
                        if (re.Read())
                        {
                            count = int.Parse(re[0].ToString());
                        }

                        if (count > 0)
                        {
                            return "帳號已存在";
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "發生錯誤";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除指定帳號
        /// </summary>
        /// <param name="account">欲刪除的帳號</param>
        /// <returns>刪除成功或失敗</returns>
        public string DelMemberByAccount(string account)
        {
            string delete = @"delete from Member where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(delete, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "刪除成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 將控制項的寬，高，左邊距，頂邊距和字體大小暫存到tag屬性中
        /// </summary>
        /// <param name="cons">遞歸控制項中的控制項</param>
        public void SetTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                    SetTag(con);
            }
        }

        //根據窗體大小調整控制項大小
        public void SetControls(float newx, float newy, Control cons)
        {
            //遍歷窗體中的控制項，重新設置控制項的值
            foreach (Control con in cons.Controls)
            {
                string[] mytag = con.Tag.ToString().Split(new char[] { ':' });//獲取控制項的Tag屬性值，並分割後存儲字元串數組
                float a = Convert.ToSingle(mytag[0]) * newx;//根據窗體縮放比例確定控制項的值，寬度
                con.Width = (int)a;//寬度
                a = Convert.ToSingle(mytag[1]) * newy;//高度
                con.Height = (int)(a);
                a = Convert.ToSingle(mytag[2]) * newx;//左邊距離
                con.Left = (int)(a);
                a = Convert.ToSingle(mytag[3]) * newy;//上邊緣距離
                con.Top = (int)(a);
                Single currentSize = Convert.ToSingle(mytag[4]) * newy;//字體大小
                con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                if (con.Controls.Count > 0)
                {
                    SetControls(newx, newy, con);
                }
            }
        }
    }
}
