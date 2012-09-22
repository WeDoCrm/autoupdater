using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using WindowsInstaller;
using System.Diagnostics;
using System.IO;

namespace UnInstallWD
{
    public partial class Form1 : Form
    {
        FileInfo fileInfo;
        public Form1()
        {
            InitializeComponent();
        }

        private void startUninstall(string productName)
        {
            try
            {
                Installer wi = (Installer)Activator.CreateInstance(Type.GetTypeFromProgID("WindowsInstaller.Installer"));
                StringList sl = wi.Products;
                Process pro = null;

                foreach (string pn in sl)
                {
                    string pc = wi.get_ProductInfo(pn, "ProductName");
                    logFileWrite(pc);
                    if (pc.Equals(productName))
                    {
                        pro = new Process();
                        pro.StartInfo.FileName = "msiexec.exe";
                        pro.StartInfo.Arguments = "/x " + pn + " /passive";
                        pro.Exited += new EventHandler(pro_Exited);
                        pro.EnableRaisingEvents = true;
                        pro.Start();
                        break;
                    }
                }
                this.Hide();
            }
            catch (Exception ex)
            {
                logFileWrite(ex.ToString());
            }
        }

        private void pro_Exited(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("WeDo 프로그램 설치제거 완료!");
            if (dr == DialogResult.OK)
            {
                Process.GetCurrentProcess().Kill();
            }
        }

        private void btn_continue_MouseClick(object sender, MouseEventArgs e)
        {
            startUninstall("WeDo Client Demo");
        }

        private void btn_cancel_MouseClick(object sender, MouseEventArgs e)
        {
            Process.GetCurrentProcess().Kill();
        }

        public void logFileWrite(string _log)
        {
            string date = DateTime.Now.ToShortDateString();
            StreamWriter sw = null;
            fileInfo = new FileInfo(Application.StartupPath + "\\Uninstaller_" + date + ".txt");
            if (!fileInfo.Exists)
            {
                fileInfo.Create();
            }
            try
            {
                sw = new StreamWriter(Application.StartupPath + "\\Uninstaller_" + date + ".txt", true);
                sw.WriteLine(_log);
                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("logFileWriter() 에러 : " + e.ToString());
            }
        }
    }
}
