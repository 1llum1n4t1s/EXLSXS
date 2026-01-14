using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace EXLSXS
{
    /// <summary>
    /// COM アドイン基盤クラス (.NET 10 対応)
    /// </summary>
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789012")]
    public class ExcelAddInBase : IDTExtensibility2
    {
        protected Application ExcelApplication;
        protected object AddInInstance;

        public ExcelAddInBase()
        {
        }

        /// <summary>IDTExtensibility2 実装: OnAddInsUpdate</summary>
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>IDTExtensibility2 実装: OnBeginShutdown</summary>
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>IDTExtensibility2 実装: OnConnection</summary>
        public void OnConnection(object application, ext_ConnectMode connectMode, object addIn, ref Array custom)
        {
            try
            {
                ExcelApplication = (Application)application;
                AddInInstance = addIn;

                OnConnected();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnConnection Error: {ex.Message}");
            }
        }

        /// <summary>IDTExtensibility2 実装: OnDisconnection</summary>
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                OnDisconnected();
                ExcelApplication = null;
                AddInInstance = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnDisconnection Error: {ex.Message}");
            }
        }

        /// <summary>OnConnection で呼ばれるコールバック</summary>
        protected virtual void OnConnected()
        {
        }

        /// <summary>OnDisconnection で呼ばれるコールバック</summary>
        protected virtual void OnDisconnected()
        {
        }
    }
}
