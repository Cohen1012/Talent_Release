using ServiceStack;
using ShareClassLibrary;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary.Model;

namespace TalentClassLibrary
{
    public class TalentSearch : SQL
    {
        private static TalentSearch talentSearch = new TalentSearch();

        public static TalentSearch GetInstance() => talentSearch;

        /// <summary>
        /// 紀錄目前發生
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 根據條件查詢符合的資料
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
        public DataTable SelectIdByFilter(string keyWords, string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate, string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Contact_Id,Name,Code_Id,CONVERT(varchar(100), Contact_Date, 111) Contact_Date,Contact_Status,Remarks,CONVERT(varchar(100), Interview_Date, 111) Interview_Date,CONVERT(varchar(100), UpdateTime, 111) UpdateTime from FilterTable where";
            try
            {
                if (Valid.GetInstance().ValidDateRange(startEditDate, endEditDate) != string.Empty)
                {
                    ErrorMessage = "最後編輯日之日期格式或者是日期區間不正確";
                    return new DataTable();
                }

                if (Valid.GetInstance().ValidDateRange(startInterviewDate, endInterviewDate) != string.Empty)
                {
                    ErrorMessage = "面談日期之日期格式或者是日期區間不正確";
                    return new DataTable();
                }

                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    this.CombinationWhereByContactFilter(places, expertises, cooperationMode, states, startEditDate, endEditDate, da);
                    this.CombinationWhereByInterviewFilter(isInterview, interviewResult, startInterviewDate, endInterviewDate, da);
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
        /// 將查詢結果組合成符合格式的DataTable
        /// </summary>
        /// <param name="dt">查詢決果</param>
        /// <returns></returns>
        public DataTable CombinationGrid(DataTable dt)
        {
            DataTable dataTable = new DataTable();

            var idList = dt.AsEnumerable().Select(x => x.Field<int>("Contact_Id")).Distinct().ToList();
            List<SearchResult> searchResultList = new List<SearchResult>();
            for (int i = 0; i < idList.Count; i++)
            {
                SearchResult searchResult = new SearchResult
                {
                    Id = idList[i].ToString(),
                };

                DataSet dataSet = new DataSet();

                DataTable contactDt = (from row in dt.AsEnumerable()
                                       where row.Field<int>("Contact_Id") == idList[i]
                                       select new
                                       {
                                           Contact_Id = row.Field<int>("Contact_Id"),
                                           Name = row.Field<string>("Name"),
                                           UpdateTime = row.Field<string>("UpdateTime")
                                       }).Distinct().LinqQueryToDataTable();
                DataTable codeDt = (from row in dt.AsEnumerable()
                                     where row.Field<int>("Contact_Id") == idList[i]
                                     select new
                                     {
                                         Code_Id = row.Field<string>("Code_Id")
                                     }).Distinct().LinqQueryToDataTable();
                DataTable statusDt = (from row in dt.AsEnumerable()
                              where row.Field<int>("Contact_Id") == idList[i]
                              select new
                              {
                                  Contact_Date = row.Field<string>("Contact_Date"),
                                  Contact_Status = row.Field<string>("Contact_Status"),
                                  Remarks = row.Field<string>("Remarks")
                              }).Distinct().OrderByDescending(x => x.Contact_Date).Take(2).LinqQueryToDataTable();
                DataTable interviewDateDt = (from row in dt.AsEnumerable() where row.Field<int>("Contact_Id") == idList[i]
                                               select new
                                               {
                                                   Interview_Date = row.Field<string>("Interview_Date")
                                               }).Distinct().OrderByDescending(x => x.Interview_Date).Take(2).LinqQueryToDataTable();   
            }
            return dataTable;
        }

        /// <summary>
        /// 組合面談資訊條件的where語法
        /// </summary>
        /// <param name="isInterview">是否已面談，值為"已面談","未面談","不限"</param>
        /// <param name="interviewResult">面談結果，值為"錄用","不錄用","暫保留","不限"</param>
        /// <param name="startInterviewDate">起始日期，日期格式</param>
        /// <param name="endInterviewDate">結束日期，日期格式</param>
        /// <param name="da"></param>
        /// <returns></returns>
        private SqlDataAdapter CombinationWhereByInterviewFilter(string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate, SqlDataAdapter da)
        {
            ////是否已面談
            if (isInterview == "已面談")
            {
                da.SelectCommand.CommandText += @" and Appointment is not null and Appointment !=''";
            }
            else if (isInterview == "未面談")
            {
                da.SelectCommand.CommandText += @" and Appointment is null or Appointment =''";
            }

            ////面談結果
            if (interviewResult != "不限")
            {
                da.SelectCommand.CommandText += @" and ISNULL(Appointment,'NA') = ISNULL(ISNULL(@interviewResult,Appointment),'NA')";
                da.SelectCommand.Parameters.Add("@interviewResult", SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty(interviewResult);
            }

            ////面談日期
            if (!string.IsNullOrEmpty(startInterviewDate) || !string.IsNullOrEmpty(endInterviewDate))
            {
                da.SelectCommand.CommandText += @" and Interview_Date <= ISNULL(@endInterviewDate, Interview_Date) AND Interview_Date >= ISNULL(@startInterviewDate, Interview_Date)";
                da.SelectCommand.Parameters.Add("@endInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endInterviewDate);
                da.SelectCommand.Parameters.Add("@startInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startInterviewDate);
            }

            return da;
        }
        /// <summary>
        /// 組合聯繫資訊的where語法
        /// </summary>
        /// <param name="places">地點，多筆請用","隔開</param>
        /// <param name="expertises">技能，多筆請用","隔開</param>
        /// <param name="cooperationMode">合作模式</param>
        /// <param name="states">聯繫狀態</param>
        /// <param name="startEditDate">起始日期，日期格式</param>
        /// <param name="endEditDate">結束日期，日期格式</param>
        /// <param name="da"></param>
        /// <returns></returns>
        private SqlDataAdapter CombinationWhereByContactFilter(string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate, SqlDataAdapter da)
        {
            string[] place = places.Split(',');
            string[] expertise = expertises.Split(',');
            da.SelectCommand.CommandText += @" ISNULL(Status,'NA') = ISNULL(ISNULL(@status,Status),'NA') and
                                                                            UpdateTime >= ISNULL(@startEditDate, UpdateTime) and
                                                                            UpdateTime <= ISNULL(@endEditDate, UpdateTime) and
                                                                            ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode,Cooperation_Mode),'NA')";
            da.SelectCommand.Parameters.Add("@CooperationMode", SqlDbType.NChar).Value = this.ValueIsAny(cooperationMode);
            da.SelectCommand.Parameters.Add("@status", SqlDbType.NChar).Value = this.ValueIsAny(states);
            da.SelectCommand.Parameters.Add("@startEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startEditDate);
            da.SelectCommand.Parameters.Add("@endEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endEditDate);
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

            return da;
        }

        /// <summary>
        /// 如果值為"不限"，NULL，空，則回傳DBNULL
        /// <param name="value">要判斷的值</param>
        /// <returns>會傳值或者是DBNULL</returns>
        private object ValueIsAny(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return DBNull.Value;
            }
            else
            {
                if (value.Equals("不限"))
                {
                    return DBNull.Value;
                }

                return value;
            }
        }
    }
}
