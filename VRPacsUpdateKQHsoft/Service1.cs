using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using Microsoft.SqlServer.Server;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;

namespace readXML
{
    public partial class Service1 : ServiceBase
    {
        public static MySqlConnection conn;
        public static MySqlDataAdapter adapter;
        public static DataSet dt = new DataSet();
        static string conStr = "";
        static string report = "";
        static string backup = "";
        static string error = "";

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            using (StreamReader r = new StreamReader(@"C:\Users\Admin\Documents\Zalo Received Files\configure.json"))
            {
                string json = r.ReadToEnd();
                List<readJson> array = JsonConvert.DeserializeObject<List<readJson>>(json);

                foreach (var item in array)
                {
                    report = item.Report_His;
                    conStr = item.Connect_DataBase;
                    backup = item.Backup_Report_His;
                    error = item.Error_Report_His;
                }
            }

            if (report == "" || conStr == "" || backup == "" || error == "")
            {
                writeLog.write("Error: file configure.json", "", "", DateTime.Now);
                return;
            }

            FileSystemWatcher watcher = new FileSystemWatcher(report);
            watcher.EnableRaisingEvents = true;
            watcher.IncludeSubdirectories = true;

            //xu ly su thay doi cua file
            watcher.Created += watcher_Created;
            Service1.readFileXML();
            Console.Read();
        }

        protected override void OnStop()
        {
        }

        public static void RunCmd(Service1 obj)
        {
            obj.OnStart(null);
            obj.OnStop();
        }

        static void readFileXML()
        {
            foreach (var file in System.IO.Directory.GetFiles(report))
            {
                ProcessFile(file);
            }
        }

        static void ProcessFile(string Filename)
        {
            try
            {
                conn = new MySqlConnection(conStr);
                //Console.Read();
                dt.ReadXml(Filename);
                string acc = "";
                string userID = "";
                string result = "";
                string proposal = "";
                string descttext = "";
                string aprotime = "";
                CultureInfo provider = new CultureInfo("en-GB");

                foreach (DataRow item in dt.Tables[0].Rows)
                {
                    acc = item["AccessionNo"].ToString();
                    userID = item["MaBacSiKetLuan"].ToString();
                    result = item["KetLuan"].ToString();
                    proposal = item["DeNghi"].ToString();
                    descttext = item["MoTa"].ToString();
                    //aprotime = item["ThoiGianThucHien"].ToString();
                    //aprotime = Convert.ToDateTime(item["ThoiGianThucHien"]).ToString("yyyy/MM/dd");
                    //aprotime = Convert.ToDateTime(item["ThoiGianThucHien"]).ToString("yyyy/MM/dd");
                    //aprotime = DateTime.ParseExact(item["ThoiGianThucHien"].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    //aprotime = Convert.ToDateTime(item["ThoiGianThucHien"].ToString(), CultureInfo.InvariantCulture);
                    //aprotime = item["ThoiGianThucHien"].ToString(format);
                    aprotime = item["ThoiGianThucHien"].ToString();
                }

                aprotime=DateTime.ParseExact(aprotime,
                                  "yyyyMMddHHmm",
                                   CultureInfo.InvariantCulture).ToString("yyyy-MM-dd HH:mm:ss");

                conn.Open();
                string queryselect = "select * from m_study where OrgCode = " + acc;
                MySqlCommand mySql = new MySqlCommand(queryselect, conn);
                var data = mySql.ExecuteReader();

                //tra ve datatable
                if (!data.HasRows)
                {
                    string sourceFile = Path.GetFileName(Filename);
                    if (File.Exists(error + "\\" + sourceFile))
                    {
                        File.Delete(error + "\\" + sourceFile);
                    }
                    File.Move(Filename, error + "\\" + sourceFile);
                }
                
                MySqlConnection conn1 = new MySqlConnection(conStr);
                conn1.Open();
                //conn.Open();
                while (data.Read())
                {
                    int id = Convert.ToInt32(data["id"]);
                    if (id > 0)
                    {
                        var query = "update m_service_ex set UserID = '" + userID + "',Result = '" + result + "',Proposal = '" + proposal + "',DescTxt ='" + descttext + "',AproTime = '" + aprotime + "',AproveByID = '" + userID + "' where StudyID ='" + id + "'";
                        MySqlCommand sqlupdate = new MySqlCommand(query, conn1);
                        
                        sqlupdate.ExecuteNonQuery();
                        
                        string sourceFile = Path.GetFileName(Filename);
                        if (File.Exists(backup + "\\" + sourceFile))
                        {
                            File.Delete(backup + "\\" + sourceFile);
                        }
                        File.Move(Filename, backup + "\\" + sourceFile);
                        conn1.Close();
                        conn1.Dispose();
                    }
                    else
                    {
                        string sourceFile = Path.GetFileName(Filename);
                        if (File.Exists(error + "\\" + sourceFile))
                        {
                            File.Delete(error + "\\" + sourceFile);
                        }
                        File.Move(Filename, error + "\\" + sourceFile);
                        conn1.Close();
                        conn1.Dispose();
                    }
                }
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                writeLog.write("Error: ProcessFile " + Filename + " " + ex.ToString(), "", "", DateTime.Now);
                conn.Close();
                conn.Dispose();
            }

        }

        private static void watcher_Renamed(object sender, RenamedEventArgs e)
        {
            //ProcessFile(e.Name);
            writeLog.write("Bạn vừa đổi tên file : ", e.Name, null, DateTime.Now);
        }

        private static void watcher_Delete(object sender, FileSystemEventArgs e)
        {
            writeLog.write("Bạn vừa xóa file : ", e.Name, null, DateTime.Now);
        }

        private static void watcher_Created(object sender, FileSystemEventArgs e)
        {
            ProcessFile(e.FullPath);
            Console.WriteLine(e.FullPath);
            writeLog.write("Create file : ", e.FullPath, null, DateTime.Now);
        }

        private static void watcher_Changer(object sender, FileSystemEventArgs e)
        {
            //ProcessFile(e.Name);
            writeLog.write("Bạn vừa thay đổi file : ", e.Name, null, DateTime.Now);
        }
    }

    public class writeLog
    {
        public static void write(string text, string name, string oldname, DateTime date)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "log.txt";

            if (System.IO.File.Exists(path))
            {
                Console.WriteLine(text + " " + name + " " + oldname + " " + DateTime.Now.ToString() + Environment.NewLine);
                System.IO.File.AppendAllText(path, text + " " + name + " " + oldname + " " + DateTime.Now.ToString() + Environment.NewLine);
            }
        }
    }

    public struct readJson
    {
        public string Connect_DataBase;
        public string Report_His;
        public string Backup_Report_His;
        public string Error_Report_His;
    }
}