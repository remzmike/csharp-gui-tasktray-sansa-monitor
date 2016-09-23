using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

// http://msdn.microsoft.com/en-us/magazine/cc164123.aspx#S6

namespace msansa
{

    [StructLayout(LayoutKind.Sequential)]
    struct DEV_BROADCAST_HDR
    {
        public uint dbch_Size;
        public uint dbch_DeviceType;
        public uint dbch_Reserved;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    struct DEV_BROADCAST_VOLUME
    {
        public uint dbcv_size;
        public uint dbcv_devicetype;
        public uint dbcv_reserved;
        public uint dbcv_unitmask;
        public UInt16 dbcv_flags;
    }

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                var letters = GetAllRemovableDriveLetters();
                CheckDriveLetters(letters);
            }
            catch (Exception ex)
            {
                DisplayException(ex);
                throw ex;
            }

        }

        // http://stackoverflow.com/questions/5931500/plug-and-play-api-for-detecting-interfacing-flash-drive-insertions
        public static int WM_DEVICECHANGE = 0x219;
        public const int DBT_DEVICEARRIVAL = 0x8000;
        public const int DBT_DEVTYP_VOLUME = 0x00000002;
        public const int DBTF_MEDIA = 0x0001;
        [DebuggerStepThrough()]
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            var P_DBT_DEVICEARRIVAL = new IntPtr(DBT_DEVICEARRIVAL);
            try
            {
                if (m.Msg == WM_DEVICECHANGE)
                {
                    debug("WM_DEVICECHANGE");
                    // http://msdn.microsoft.com/en-us/library/aa363205.aspx
                    if (m.WParam == P_DBT_DEVICEARRIVAL)
                    {
                        debug("DBT_DEVICEARRIVAL");

                        // http://www.tech-archive.net/Archive/DotNet/microsoft.public.dotnet.framework.interop/2007-01/msg00224.html
                        DEV_BROADCAST_HDR hdr = (DEV_BROADCAST_HDR)m.GetLParam(typeof(DEV_BROADCAST_HDR));
                        if (hdr.dbch_DeviceType == DBT_DEVTYP_VOLUME)
                        {
                            debug("DBT_DEVTYP_VOLUME");
                            DEV_BROADCAST_VOLUME vol = (DEV_BROADCAST_VOLUME)m.GetLParam(typeof(DEV_BROADCAST_VOLUME));
                            //debug(String.Format("vol.dbcv_flags & DBTF_MEDIA == {0}", vol.dbcv_flags & DBTF_MEDIA));
                            if ((vol.dbcv_flags & DBTF_MEDIA) == 0)
                            {
                                debug("DBTF_MEDIA");
                                debug(String.Format("vol.dbcv_unitmask : {0}", vol.dbcv_unitmask));
                                var unitmask = vol.dbcv_unitmask;
                                var letters = GetDriveLettersFromMask(unitmask);
                                CheckDriveLetters(letters);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayException(ex);
                throw ex;
            }

            base.WndProc(ref m);
        }

        public void debug(string msg)
        {
            if (true)
            {
                // screw richtext coloring
                var timestamp = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                output.AppendText(timestamp);
                output.AppendText(" : " + msg);
                output.AppendText(Environment.NewLine);
            }
        }

        public void DisplayException(Exception ex)
        {
            debug("EXCEPTION : ");
            debug(ex.Message);
            debug(ex.StackTrace);
        }

        public List<char> GetDriveLettersFromMask(UInt32 unitmask)
        {
            var result = new List<char>();

            for (var i = 0; i < 26; i++)
            {
                if ((unitmask & 0x1) == 1)
                {
                    var letter = (char)(i + 65);
                    result.Add(letter);
                }
                unitmask = unitmask >> 1;
            }

            return result;
        }

        public List<char> GetAllRemovableDriveLetters()
        {
            DriveInfo[] drives = DriveInfo.GetDrives();
            var result = new List<char>();

            foreach (DriveInfo drive in drives)
            {
                if (drive.IsReady)
                {
                    if (drive.DriveType == DriveType.Removable)
                    {
                        result.Add(drive.Name[0]);
                    }
                    /*
                    double fspc = 0.0;
                    double tspc = 0.0;
                    double percent = 0.0;

                    fspc = drive.TotalFreeSpace;
                    tspc = drive.TotalSize;
                    percent = (fspc / tspc) * 100;
                    float num = (float)percent;

                    debug(String.Format("Drive:{0} With {1} % free", drive.Name, num));
                    debug(String.Format("Space Reamining:{0}", drive.AvailableFreeSpace));
                    debug(String.Format("Percent Free Space:{0}", percent));
                    debug(String.Format("Space used:{0}", drive.TotalSize));
                    debug(String.Format("Type: {0}", drive.DriveType));
                    */
                }
            }
            return result;
        }

        public void CheckDriveLetters(List<char> letters)
        {
            foreach (var letter in letters)
            {
                debug(String.Format("checking drive {0}", letter));
                if (IsSansaDrive(letter))
                {
                    debug(String.Format("sansa drive {0} detected", letter));
                    HandleSansaDrive(letter);
                }
                else
                {
                    //debug(String.Format("not a sansa drive: {0}", letter));
                }
            }        
        }

        public bool IsSansaDrive(char letter)
        {
            var fpath = String.Format(@"{0}:\MTABLE.SYS", letter);
            //var fpath = String.Format(@"{0}:\MTABLE.SYS"); // to test exceptions
            var result = File.Exists(fpath);
            debug(String.Format("is sansa, {0}, {1}", fpath, result));
            return result;
        }

        public void HandleSansaDrive(char letter)
        {            
            var now = DateTime.Now;
	        var format = "yyyy.MM.dd-HH.mm.ss-fffffff";
	        string timestamp = now.ToString(format);

            string modelName = GetModelName(letter);
            string droot = String.Format(@"C:\sansa.{0}", modelName);

            Directory.CreateDirectory(droot);

            var ipath1 = String.Format(@"{0}:\MTABLE.SYS", letter);
            var opath1 = String.Format(@"{0}\{1}.MTABLE.SYS", droot, timestamp);
            CopyFile(ipath1, opath1);

            if (modelName!="clip") // clipplus and clipzip use both files
            {
                var ipath2 = String.Format(@"{0}:\RES_INFO.SYS", letter);
                var opath2 = String.Format(@"{0}\{1}.RES_INFO.SYS", droot, timestamp);
                CopyFile(ipath2, opath2);
            }
        }

        public void CopyFile(string ipath, string opath, bool gzip = true)
        {           
            File.Copy(ipath, opath);
            if (gzip) GZipFile(opath);
        }

        public string GetModelName(char letter)
        {
            // both clip zip & clip+ have this DID.bin file
            // the regular clip does not
            //string path = String.Format(@"{0}:\DID.bin", letter);
            //return File.Exists(path);
            // all of them have the version.sdk file though
            const string clip = "Clip";
            const string clipzip = "Clip Zip";
            const string clipplus = "Clip+";
            
            TextReader tr = new StreamReader(letter + ":\\version.sdk");
            tr.ReadLine();
            tr.ReadLine();
            string versionLine = tr.ReadLine();
            string modelName = versionLine.Split(':')[1].Trim();
            switch (modelName)
            {
                case clip:
                    return "clip";
                case clipplus:
                    return "clipplus";
                case clipzip:
                    return "clipzip";
                default:
                    throw new Exception("unexpected model string: {0}");
            }
        }

        // gzip the file at this path, then delete original on success
        public void GZipFile(string fpath)
        {                  
            byte[] data = File.ReadAllBytes(fpath);
            var gzpath = fpath + ".gz";
            debug("gzipping to " + gzpath);
            GZipToFile(gzpath, data);
            debug("deleting " + fpath);
            File.Delete(fpath);
        }

        private void GZipToFile(string fpath, byte[] data)
        {            
            using (FileStream fs = new FileStream(fpath, FileMode.Create))
            using (GZipStream gz = new GZipStream(fs, CompressionMode.Compress, false))
            {
                gz.Write(data, 0, data.Length);
            }            
        }

        // ui garbage

        // http://stackoverflow.com/questions/3563889/how-to-let-windows-form-exit-to-system-tray
        // http://stackoverflow.com/questions/3571477/why-is-my-notifyicon-application-not-exiting-when-i-use-application-exit-or-fo
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // false when we call Application.Exit
            // (true if we call this.Close or user closes)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
        }
        
        void RestoreWindow()
        {
            this.WindowState = FormWindowState.Normal;
            this.Show();
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            RestoreWindow();
        }

        private void openMsansaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RestoreWindow();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
