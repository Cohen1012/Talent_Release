using ShareClassLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalentClassLibrary
{
    /// <summary>
    /// 檔案處理類別
    /// </summary>
    public class TalentFiles
    {
        private static TalentFiles talentFiles = new TalentFiles();

        public static TalentFiles GetInstance() => talentFiles;

        /// <summary>
        /// 下載附加檔案
        /// </summary>
        /// <param name="serverPath">伺服器端的檔案路徑</param>
        /// <param name="clientPath">存檔的路徑</param>
        /// <returns></returns>
        public string DownloadFile(string serverPath, string clientPath)
        {
            string msg = Valid.GetInstance().ValidFilePath(serverPath);
            if (msg != string.Empty)
            {
                return "不存在的路徑";
            }

            try
            {
                using (Stream stream = File.Open(serverPath, FileMode.Open))
                {
                    using (FileStream fs = new FileStream(clientPath, FileMode.OpenOrCreate))
                    {
                        stream.CopyTo(fs);
                        fs.Flush();
                    }
                }

                return "下載成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "下載失敗";
            }

        }

        /// <summary>
        /// 上傳照片
        /// </summary>
        /// <param name="path">本機端路徑</param>
        /// <param name="interviewId">面談ID</param>
        /// <returns></returns>
        public string UpLoadImage(string path, string interviewId)
        {
            string msg = Valid.GetInstance().ValidFilePath(path);
            if (msg != string.Empty)
            {
                return "不存在的路徑";
            }

            try
            {
                ////取得副檔名
                string extension = Path.GetExtension(path).ToLowerInvariant();
                // 檢查 Server 上該資料夾是否存在，不存在就自動建立
                string serverDir = @".\images";
                if (Directory.Exists(serverDir) == false)
                {
                    Directory.CreateDirectory(serverDir);
                }

                ////將檔名以面談ID命名
                string fileName = string.Format("{0}{1}", interviewId, extension);
                ////組合出完整路徑
                string serverFilePath = Path.Combine(serverDir, fileName);

                using (Stream stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    using (FileStream fs = new FileStream(serverFilePath, FileMode.OpenOrCreate))
                    {
                        stream.CopyTo(fs);
                        fs.Flush();
                    }
                }
                return serverFilePath;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "上傳失敗";
            }
        }

        /// <summary>
        /// 檢查目錄的檔案與UI的檔案，如果目錄的檔案沒出現在UI的檔案中就刪除之
        /// </summary>
        /// <param name="files">UI檔案</param>
        /// <param name="serverDir">DB檔案存在的目錄</param>
        private void CheckDBFileIsExists(List<string> files, string serverDir)
        {
            foreach (string fileNameDB in Directory.GetFiles(serverDir))
            {
                bool isExists = false;
                foreach (string path in files)
                {
                    string fileNameUI = Path.GetFileName(path);
                    if (fileNameUI.Equals(Path.GetFileName(fileNameDB)))
                    {
                        isExists = true;
                        break;
                    }
                }

                if (isExists == false)
                {
                    File.Delete(fileNameDB);
                }
            }
        }

        /// <summary>
        /// 將附加檔案上傳到Server
        /// </summary>
        /// <param name="files">附加檔案</param>
        /// <param name="interviewId">面談ID</param>
        /// <param name="attachedFileMode">附加檔案模式</param>
        /// <returns></returns>
        public List<string> SaveAttachedFiles(List<string> files, string interviewId, string attachedFileMode)
        {
            List<string> serverFilePathList = new List<string>();
            try
            {
                //// 檢查 Server 上該資料夾是否存在，不存在就建立
                string serverDir = @".\files\" + interviewId + @"\" + attachedFileMode;
                if (Directory.Exists(serverDir) == false)
                {
                    Directory.CreateDirectory(serverDir);
                }

                CheckDBFileIsExists(files, serverDir);

                foreach (string path in files)
                {
                    ////組合出完整路徑
                    string serverFilePath = Path.Combine(serverDir, Path.GetFileName(path));
                    ////目錄相同代表相同檔案，不做事
                    if (path.Equals(serverFilePath))
                    {
                        serverFilePathList.Add(serverFilePath);
                        continue;
                    }

                    using (Stream stream = File.Open(path, FileMode.Open, FileAccess.Read))
                    {
                        using (FileStream fs = new FileStream(serverFilePath, FileMode.OpenOrCreate))
                        {
                            stream.CopyTo(fs);
                            fs.Flush();
                        }
                    }

                    serverFilePathList.Add(serverFilePath);
                }

                return serverFilePathList;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                serverFilePathList.Clear();
                serverFilePathList.Add("上傳失敗");
                return serverFilePathList;
            }
        }

        /// <summary>
        /// 刪除指定面談資料的附加檔案
        /// </summary>
        /// <param name="interviewId"></param>
        /// <returns></returns>
        public string DelFilesByInterviewId(string interviewId)
        {
            string serverDir = @".\files\" + interviewId;
            try
            {
                if (Directory.Exists(serverDir))
                {
                    Directory.Delete(serverDir, true);
                }

                return "附加檔案刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "附加檔案刪除失敗";
            }
        }

        /// <summary>
        /// 根據情況處理圖片檔案
        /// </summary>
        /// <param name="dbPath"></param>
        /// <param name="uiPath"></param>
        /// <param name="interviewId"></param>
        public string UpdateImage(string dbPath, string uiPath, string interviewId)
        {
            ////DB跟UI都沒有圖片則不做事
            if (string.IsNullOrEmpty(dbPath) && string.IsNullOrEmpty(uiPath))
            {
                return string.Empty;
            }

            ////如果照片沒有換
            if (dbPath.Equals(uiPath))
            {
                return dbPath;
            }

            ////DB有圖片，UI沒圖片則砍掉圖片檔案
            if (!string.IsNullOrEmpty(dbPath) && string.IsNullOrEmpty(uiPath))
            {
                try
                {
                    if (File.Exists(dbPath))
                    {
                        File.Delete(dbPath);
                    }

                    return uiPath;
                }
                catch (IOException ex)
                {
                    LogInfo.WriteErrorInfo(ex);
                    return "檔案刪除失敗";
                }
            }

            return this.UpLoadImage(uiPath, interviewId);
        }
    }
}
