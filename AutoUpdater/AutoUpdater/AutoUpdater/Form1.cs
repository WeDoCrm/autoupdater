using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Xml;
using System.Diagnostics;
using WindowsInstaller;
using System.Threading;

namespace AutoUpdater
{
    public partial class Form1 : Form
    {
        string FtpHost_uri = null;
        string ftpPass = "eclues!@";
        string ftpid = "eclues";
        string downfilename = null;
        string exeDir = null;
        string localDir = null;
        string productName = null;
        string backupFileName = null;
        string SVRver = null;
        string id = null;
        string extension = null;
        string pass = null;
        string autostart = null;
        string topmost = null;
        string save_pass = null;
        string server_addr = null;
        string nopop = null;
        WebClient wc = null;
        Process pro = null;

        delegate void Labeldelegate(string msg);

        public Form1()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            //Thread thread = new Thread(new ThreadStart(dummyThread));
            //thread.Start();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason != CloseReason.TaskManagerClosing)
            {
                e.Cancel = true;
            }
        }

        public void dummyThread()
        {
            while (true)
            {
                Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                loadConfigXml();

                Uri ftpuri = new Uri(FtpHost_uri);
                FtpWebRequest wr = (FtpWebRequest)WebRequest.Create(ftpuri);
                wr.Method = WebRequestMethods.Ftp.ListDirectory;
                wr.Credentials = new NetworkCredential(ftpid, ftpPass);
                FtpWebResponse wres = (FtpWebResponse)wr.GetResponse();
                Stream st = wres.GetResponseStream();
                SVRver = null;
                bool isUpdate = false;
                if (st.CanRead)
                {
                    StreamReader sr = new StreamReader(st);
                    SVRver = sr.ReadLine();
                    FtpHost_uri += SVRver + "/" + downfilename;
                }

                logFileWrite("ftpuri : " + FtpHost_uri);

                DelDir();

                Thread.Sleep(3000);
                wc = new WebClient();
                ftpuri = new Uri(FtpHost_uri);
                logFileWrite("Download dir : " + ftpuri.ToString());
                wc.Credentials = new NetworkCredential(ftpid, ftpPass);
                //wc.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.CacheIfAvailable);
                // 다운로드 결과 및 진행 상황을 위한 핸들러 설정 DownloadFileAsync()를 이용할 경우에만 동작함.
                wc.DownloadFileCompleted += new AsyncCompletedEventHandler(wc_DownloadFileCompleted);
                wc.DownloadProgressChanged += new DownloadProgressChangedEventHandler(wc_DownloadProgressChanged);
                // 파일 다운로드 DownloadFile()을 이용할 경우 DownloadFileComplited, DownloadProgressChanged 이벤트를 사용할 수 없음.
                wc.DownloadFileAsync(ftpuri, @"C:\temp\" + downfilename);

            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString() + "_" + DateTime.Now.ToString());
            }
        }

        private void DelDir()
        {
            //DirectoryInfo di = new DirectoryInfo(@"C:\"+DateTime.Now.ToShortDateString());
            DirectoryInfo di = new DirectoryInfo(@"C:\temp");

            if (!di.Exists)
            {
                di.Create();
            }
            else
            {
                //FileInfo fi = new FileInfo("C:\\" + DateTime.Now.ToShortDateString() + "\\" + downfilename);

                FileInfo fi = new FileInfo("C:\\temp\\" + downfilename);

                if (fi.Exists)
                {
                    fi.Delete();
                }
                //if (fi.Exists)
                //{
                //    fi.Delete();
                //    fi.Refresh();
                //}
            }
        }

        private void wc_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            userInfoBackup();
            Uninstall();
            Labeldelegate dele = new Labeldelegate(changeLabelText);
            Invoke(dele, new object[]{"WeDo 업데이트중..."});
        }

        

        private void Uninstall()
        {
            try
            {
                Labeldelegate dele = new Labeldelegate(changeLabelText);
                Invoke(dele, new object[] { "기존프로그램 삭제중" });

                //File.Copy(localDir + backupFileName, "C:\\temp\\" + backupFileName, true);

                Installer wi = (Installer)Activator.CreateInstance(Type.GetTypeFromProgID("WindowsInstaller.Installer"));
                StringList sl = wi.Products;

                foreach (string pn in sl)
                {
                    string pc = wi.get_ProductInfo(pn, "ProductName");
                    if (pc.Equals(productName))
                    {
                        pro = new Process();
                        pro.StartInfo.FileName = "msiexec.exe";
                        pro.StartInfo.Arguments = "/x " + pn + " /passive";
                        pro.Exited += new EventHandler(uninstall_Exited);
                        pro.EnableRaisingEvents = true;
                        pro.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }

        private void uninstall_Exited(object sender, EventArgs e)
        {
            reInstall();
        }

        private void reInstall()
        {
            try
            {
                logFileWrite("reInstall Start");
                pro = new Process();
                pro.StartInfo.FileName = "msiexec.exe";
                pro.StartInfo.Arguments = "/i C:\\temp\\" + downfilename + " /passive";
                pro.Exited += new EventHandler(reinstall_Exited);
                pro.EnableRaisingEvents = true;
                pro.Start();
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }

        private void reinstall_Exited(object sender, EventArgs e)
        {
            try
            {
                //File.Copy("C:\\temp\\" + backupFileName, localDir + backupFileName, true);
                System.Threading.Thread.Sleep(2000);
                logFileWrite(productName+" restart!!");
                versionChange();
                Process.Start(exeDir);
                this.Hide();
                MessageBox.Show("WeDo 업데이트가 완료되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Thread.Sleep(2000);
                Process.GetCurrentProcess().Kill();
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }


        private void changeLabelText(string msg)
        {
            label1.Text = msg;
            //label1.Update();
        }

        private void wc_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            // 프로그레스바에 다운로드 상태 표시
            label1.Text = "업데이트 파일 다운로드 중...";
            progressBar1.Value = e.ProgressPercentage;
        }
        

        public void logFileWrite(string _log)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(Application.StartupPath + "\\AutoUpdater_log_" + System.DateTime.Now.ToShortDateString() + ".txt", true);
                sw.WriteLine(_log);
                sw.Close();
            }
            catch (Exception e)
            {
                //Console.WriteLine("logFileWriter() 에러 : " + e.ToString());
            }
        }

        /// <summary>
        /// Ftp 관련 설정파일 값을 app.config 파일에서 로드함.
        /// </summary>
        private void loadConfigXml()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(Application.StartupPath+"\\AutoUpdater.exe.config");

                XmlNode pnode = doc.SelectSingleNode("//appSettings");
                if (pnode.HasChildNodes)
                {
                    XmlNodeList nodelist = pnode.ChildNodes;
                    foreach (XmlNode node in nodelist)
                    {
                        if (node.Attributes["key"].Value.Equals("FtpPass"))
                        {
                            //ftpPass = node.Attributes["value"].Value;
                            //logFileWrite("ftpPass : " + ftpPass);
                        }
                        else if (node.Attributes["key"].Value.Equals("FtpUserName"))
                        {
                            //ftpid = node.Attributes["value"].Value;
                            //logFileWrite("ftp_username : " + ftpid);
                        }

                        else if (node.Attributes["key"].Value.Equals("FtpHost"))
                        {
                            FtpHost_uri = node.Attributes["value"].Value;
                            logFileWrite("FtpHost_uri : " + FtpHost_uri);
                        }

                        else if (node.Attributes["key"].Value.Equals("FtpFileName"))
                        {
                            downfilename = node.Attributes["value"].Value;
                            logFileWrite("FtpFileName : " + downfilename);
                        }

                        else if (node.Attributes["key"].Value.Equals("EXEDir"))
                        {
                            exeDir = node.Attributes["value"].Value;
                            logFileWrite("EXEDir : " + exeDir);
                        }
                        else if (node.Attributes["key"].Value.Equals("ProductName"))
                        {
                            productName = node.Attributes["value"].Value;
                            logFileWrite("ProductName : " + productName);
                        }
                        else if (node.Attributes["key"].Value.Equals("LocalDir"))
                        {
                            localDir = node.Attributes["value"].Value;
                            logFileWrite("LocalDir : " + localDir);
                        }
                        else if (node.Attributes["key"].Value.Equals("backupFileName"))
                        {
                            backupFileName = node.Attributes["value"].Value;
                            logFileWrite("backupFileName : " + backupFileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }

        private void versionChange()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(localDir + "\\WDMsg_Client.exe.config");

                XmlNode pnode = doc.SelectSingleNode("//appSettings");
                if (pnode.HasChildNodes)
                {
                    XmlNodeList nodelist = pnode.ChildNodes;
                    foreach (XmlNode node in nodelist)
                    {
                        if (node.Attributes["key"].Value.Equals("FtpVersion"))
                        {
                            node.Attributes["value"].Value = SVRver;
                            logFileWrite("Version Change complete! = " + SVRver);
                        }

                        if (node.Attributes["key"].Value.Equals("id"))
                        {
                            node.Attributes["value"].Value = id;
                        }

                        if (node.Attributes["key"].Value.Equals("extension"))
                        {
                            node.Attributes["value"].Value = extension;
                        }

                        if (node.Attributes["key"].Value.Equals("pass"))
                        {
                            node.Attributes["value"].Value = pass;
                        }

                        if (node.Attributes["key"].Value.Equals("autostart"))
                        {
                            node.Attributes["value"].Value = autostart;
                        }

                        if (node.Attributes["key"].Value.Equals("topmost"))
                        {
                            node.Attributes["value"].Value = topmost;
                        }

                        if (node.Attributes["key"].Value.Equals("save_pass"))
                        {
                            node.Attributes["value"].Value = save_pass;
                        }

                        if (node.Attributes["key"].Value.Equals("serverip"))
                        {
                            node.Attributes["value"].Value = server_addr;
                        }

                        if (node.Attributes["key"].Value.Equals("nopop"))
                        {
                            node.Attributes["value"].Value = nopop;
                        }
                    }
                }

                doc.Save(localDir + "\\WDMsg_Client.exe.config");
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }

        private void userInfoBackup()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                try
                {
                    doc.Load(localDir + "\\WDMsg_Client.exe.config");
                }
                catch (FileNotFoundException fnfe1)
                {
                    try
                    {
                        logFileWrite(localDir + "\\WDMsg_Client.exe.config  파일이 없습니다.");
                        doc.Load("c:\\Program Files\\eclues\\WeDo\\WDMsg_Client.exe.config");
                    }
                    catch (FileNotFoundException fnfe2)
                    {
                        logFileWrite("c:\\Program Files\\eclues\\WeDo\\WDMsg_Client.exe.config  파일이 없습니다.");
                        doc.Load("c:\\Program Files(x86)\\eclues\\WeDo\\WDMsg_Client.exe.config");
                    }
                }
              
                XmlNode pnode = doc.SelectSingleNode("//appSettings");
                if (pnode.HasChildNodes)
                {
                    XmlNodeList nodelist = pnode.ChildNodes;
                    foreach (XmlNode node in nodelist)
                    {
                        if (node.Attributes["key"].Value.Equals("id"))
                        {
                            id = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("extension"))
                        {
                            extension = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("pass"))
                        {
                            pass = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("autostart"))
                        {
                            autostart = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("topmost"))
                        {
                            topmost = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("save_pass"))
                        {
                            save_pass = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("serverip"))
                        {
                            server_addr = node.Attributes["value"].Value;
                        }

                        if (node.Attributes["key"].Value.Equals("nopop"))
                        {
                            nopop = node.Attributes["value"].Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }
    }
}
