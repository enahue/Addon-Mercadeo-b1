using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

namespace SBOPlugins.Enumerations
{
    public enum eFileDialog { en_OpenFile = 0, en_SaveFile = 1 };
}
namespace UI_API_Csharp
{

    public class GetFileNameClass : IDisposable
    {
        #region The class implements FileDialog for open in front of B1 window

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        System.Windows.Forms.FileDialog _oFileDialog;

        // Properties
        public string FileName
        {
            get { return _oFileDialog.FileName; }
            set { _oFileDialog.FileName = value; }
        }

        public string[] FileNames
        {
            get { return _oFileDialog.FileNames; }
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

        //// Constructor
        //public GetFileNameClass()
        //{
        //    _oFileDialog = new OpenFileDialog();
        //}

        // Constructor
        public GetFileNameClass(SBOPlugins.Enumerations.eFileDialog dlg)
        {
            switch ((int)dlg)
            {
                case 0: _oFileDialog = new System.Windows.Forms.OpenFileDialog(); break;
                case 1: _oFileDialog = new System.Windows.Forms.SaveFileDialog(); break;
                default: throw new ApplicationException("GetFileNameClass Incorrect Parameter");
            }
        }

        public GetFileNameClass()
            : this(SBOPlugins.Enumerations.eFileDialog.en_OpenFile)
        {

        }

        // Dispose
        public void Dispose()
        {
            _oFileDialog.Dispose();
        }

        // Methods

        public void GetFileName()
        {
            IntPtr ptr = GetForegroundWindow();

            WindowWrapper oWindow = new WindowWrapper(ptr);

            if (_oFileDialog.ShowDialog(oWindow) != System.Windows.Forms.DialogResult.OK)
            {
                _oFileDialog.FileName = string.Empty;
            }
            oWindow = null;
        } // End of GetFileName

        #endregion

        #region WindowWrapper : System.Windows.Forms.IWin32Window

        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private IntPtr _hwnd;

            // Property
            public virtual IntPtr Handle
            {
                get { return _hwnd; }
            }

            // Constructor
            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }
        }
        #endregion
    }
}
