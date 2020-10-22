using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    internal class Util
    {
        private RichTextBox g_richTB_Status = null;
        private FileStream FWriter = null;
        private StreamWriter SWriter = null;
        private string path = "";

        public Util(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;

            var Dir = System.IO.Directory.GetCurrentDirectory();
            path = Dir + "\\KPM_Log.txt";
            FWriter = new FileStream(path, FileMode.Create, FileAccess.Write);
            SWriter = new StreamWriter(FWriter, Encoding.Default);
            DateTime current = DateTime.Now;
            DebugPrint(current.ToString("yyyy/MM/dd HH:mm:ss") + ": Please Select Excel File.");
        }

        public void DebugPrint(string sDebugString)
        {
            g_richTB_Status.Text = sDebugString;
            g_richTB_Status.Update();
            System.Diagnostics.Debug.WriteLine(sDebugString);

            if (FWriter == null)
            {
                FWriter = new FileStream(path, FileMode.Append, FileAccess.Write);
                SWriter = new StreamWriter(FWriter, Encoding.Default);
            }
            SWriter.WriteLine(sDebugString);

            SWriter.Close();
            SWriter = null;
            FWriter.Close();
            FWriter = null;
        }
    }
}