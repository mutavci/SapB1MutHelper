using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace SapB1MutHelper
{
    public class AttachmentCreate
    {
        private readonly OpenFileDialog _oFileDialog;

        public byte[] FileArray;

        //public string FileName;   
        public string FilePath;
        public string PathName;

        public string aa;

        public AttachmentCreate()
        {
            _oFileDialog = new OpenFileDialog();
        }

        // Properties
        public string FileName
        {
            get { return _oFileDialog.FileName; }
            set { _oFileDialog.FileName = value; }
        }

        public string NameFile 
        {
            get { return _oFileDialog.FileName; }  // oFileDialog.SafeFileName;
        }

        public string Filter
        {
            get { return _oFileDialog.Filter; }
            set { _oFileDialog.Filter = value; }
        }

        public string InitialDirectory
        {
            get { return _oFileDialog.InitialDirectory; }
            set { _oFileDialog.InitialDirectory = value; }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();


        public string ShowOpenFileDialog(bool openFolder = false)
        {
            var oGetFileName = new AttachmentCreate();
            oGetFileName.Filter = "All files (*.*)|*.*";
            oGetFileName.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            var threadGetExcelFile = new Thread(oGetFileName.ShowFolderBrowser);
            threadGetExcelFile.SetApartmentState(ApartmentState.STA);
            try
            {
                threadGetExcelFile.Start();
                while (!threadGetExcelFile.IsAlive)
                {
                }

                Thread.Sleep(1);
                threadGetExcelFile.Join();

                if (!string.IsNullOrEmpty(oGetFileName.FileName))
                {
                    PathName = oGetFileName.NameFile;
                    return oGetFileName.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return "";
        }


        public void ShowFolderBrowser()
        {

            var ptr = GetForegroundWindow();
            var oWindow = new WindowWrapper(ptr);
            if (_oFileDialog.ShowDialog(oWindow) != DialogResult.OK) 
                FileName = _oFileDialog.FileName;
            else
                Application.ExitThread();
        }


        public void SaveFiles(byte[] fs)
        {
            var saveFileDialog1 = new SaveFileDialog();
            //saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif";
            saveFileDialog1.Title = "Kaydedilecek yer seçiniz";
            saveFileDialog1.ShowDialog();


            if (saveFileDialog1.FileName != "")
            {

                if (!Directory.Exists(saveFileDialog1.InitialDirectory + "\\" + saveFileDialog1.FileName))
                    Directory.CreateDirectory(saveFileDialog1.InitialDirectory + "\\" + saveFileDialog1.FileName);

                var memoryStream = new MemoryStream(fs);
                var fileStream = new FileStream(saveFileDialog1.InitialDirectory + "\\" + saveFileDialog1.FileName,
                    FileMode.CreateNew);
                memoryStream.WriteTo(fileStream);
                memoryStream.Close();
                fileStream.Close();
            }
        }
    }
}