using DB.Utils;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Timers;

namespace MSTSCWindowsService
{
    public partial class Service1 : ServiceBase
    {
        private Process p = null;
        private Timer timer = new Timer();
        private byte MSTSCStatus = 0;
        private long currentID = 0;
        public Service1()
        {
            InitializeComponent();
            p = new Process();
            p.StartInfo.FileName = "cmd.exe";//设置启动的应用程序
            p.StartInfo.UseShellExecute = false;//禁止使用操作系统外壳程序启动进程
            p.StartInfo.RedirectStandardInput = true;//应用程序的输入从流中读取
            p.StartInfo.RedirectStandardOutput = true;//应用程序的输出写入流中
            p.StartInfo.RedirectStandardError = true;//将错误信息写入流
            p.StartInfo.CreateNoWindow = true;//是否在新窗口中启动进程
            timer.Interval = 1000;
            timer.Elapsed += Timer_Elapsed;
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            PortListen();
        }

        protected override void OnStart(string[] args)
        {
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter("C:\\MSTSClog.txt", true))
            {
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Start.");
            }
            p.Start();
            timer.Start();
        }

        protected override void OnStop()
        {
            using (System.IO.StreamWriter sw = new System.IO.StreamWriter("C:\\MSTSClog.txt", true))
            {
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + "Stop.");
            }
            p.Close();
            timer.Stop();
        }

        private void PortListen()
        {
            //p.StandardInput.WriteLine(@"cls");
            //p.StandardInput.WriteLine(@"netstat -a -n>c:\port.txt");//将字符串写入文本流
            p.StandardInput.WriteLine(@"netstat -n -p tcp | find "":3389""");
            string str;
            System.Threading.Thread.Sleep(100);
            while (!string.IsNullOrEmpty((str = p.StandardOutput.ReadLine())))
            {
                Console.WriteLine(str);
                if (str.Contains("TCP") && str.Contains("ESTABLISHED"))
                {
                    if (MSTSCStatus >= 1)
                    {
                        MSTSCStatus = 1;
                        return;
                    }
                    MatchCollection matches = Regex.Matches(str, @"\d+\.\d+\.\d+\.\d+");
                    if (matches.Count > 1 && matches[0].Success)
                    {
                        currentID = InsertRecord(matches[0].Value, matches[1].Value, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                    MSTSCStatus = 1;
                }
                else if (MSTSCStatus >= 1 && MSTSCStatus <= 6)
                {
                    MSTSCStatus++;
                }
                else if (MSTSCStatus > 6)
                {
                    MSTSCStatus = 0;
                    UpdateRecord(DateTime.Now.ToString(), currentID);
                }
            }
        }

        private long InsertRecord(string hostIP, string clietIP,string startTime)
        {
            long recordid = 0;
            string retidstring = "Select MSTSCinfo_autoinc_seq.nextval from dual";
            Object result = OracleHelper.GetSingle(retidstring, null);
            bool convertresult = long.TryParse(result.ToString(), out recordid);
            if (!convertresult)
            {
                return 0;
            }
            string insertstring = "insert into MSTSCInfo(ID,HostIP,ClientIP,StartTime) values(:id,:hostIP,:clientIP,to_date(:startTime,'yyyy-MM-dd hh24:mi:ss'))";
            int insertresult = OracleHelper.ExecuteNonQuery(System.Data.CommandType.Text, insertstring, new OracleParameter[]
            {
                new OracleParameter(":id", recordid),
                new OracleParameter(":hostIP", hostIP),
                new OracleParameter(":clientIP", clietIP),
                new OracleParameter(":startTime",startTime)
            });
            if (insertresult < 1)
            {
                return 0;
            }
            return recordid;
        }

        private int UpdateRecord(string endTime, long id)
        {
            string updatestring = "Update MSTSCInfo Set EndTime = to_date(:endtime,'yyyy-MM-dd hh24:mi:ss') Where ID = :id";
            int updateresult = OracleHelper.ExecuteNonQuery(System.Data.CommandType.Text, updatestring, new OracleParameter[]
            {
                new OracleParameter(":endtime", endTime),
                new OracleParameter(":id", id)
            });
            return updateresult;
        }
    }
}
