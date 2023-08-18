using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Empire.Solidworks.NewApp
{
    /// <summary>
    /// Our SolidWorks taskpane add-in
    /// </summary>
    public class TaskpaneIntegration : SwAddin
    {
        #region Privat Members

        /// <summary>
        /// The cookie to the current instance of SolidWorks we are running inside of
        /// </summary>
        private int mSwCookie;

        /// <summary>
        /// The taskpane vewi for our add-in
        /// </summary>
        private TaskpaneView mTaskpaneView;

        /// <summary>
        /// The UI control that is going to be inside the Solidworks taskpane view
        /// </summary>

        private TaskpaneHostUI mTaskpaneHost;

        /// <summary>
        /// The current instance of the SolidWorks application
        /// </summary>
        private SldWorks mSolidworksApplication;


        #endregion

        #region Public Members
        /// <summary>
        /// The unique Id to the taskpane used for registeration in COM
        /// </summary>
        public const string SWTASKPANE_PROGID = "Empire.Solidworks.NewApp.Taskpane";

        #endregion

        #region Solidworks Add-in Callbacks

        /// <summary>
        /// Called when SolidWorks has loaded our add-in and wants us to do our connection logic
        /// </summary>
        /// <param name="ThisSW">The current SolidWorks instance</param>
        /// <param name="Cookie">The current SolidWorks cookie Id</param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // Store a reference to the current SolidWorks instance
            mSolidworksApplication = (SldWorks)ThisSW;

            // Store cookie Id
            mSwCookie = Cookie;

            // Setup Callback info
            var ok = mSolidworksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);

            // Create our UI
            LoadUi();

            // Return OK
            return true;

        }


        /// <summary>
        /// Called when SolidWorks is about to unload our add-in and wants us to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public bool DisconnectFromSW()
        {
            // Clean up our UI
            UnloadUI();

            // Return OK
            return true;
        }

        #endregion

        #region Create UI

        /// <summary>
        /// Create our Taskpane and inject our host UI
        /// </summary>
        private void LoadUi()
        {
            // Find location to our taskpan icon
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "SolidworksAppBrain.png");

            // Create our Taskpane
            mSolidworksApplication.CreateTaskpaneView2(imagePath, "Woo, My first SqAddin");
            // @"C:\Program Files\Repo\SolidworksAPICodes\Referances\SolidworksAppBrain.png" in case we need to hard code the location

            // Load our UI into the taskpane
            mTaskpaneHost = (TaskpaneHostUI)mTaskpaneView.AddControl(TaskpaneIntegration.SWTASKPANE_PROGID, string.Empty);
        }

        private void UnloadUI()
        {
            mTaskpaneHost = null;

            //Remove taskpane view
            mTaskpaneView.DeleteView();

            // Release COM Reference and cleanup memory
            Marshal.ReleaseComObject(mTaskpaneView);

            mTaskpaneView = null;
        }

        #endregion

        #region COM Registration

        /// <summary>
        /// The COM Registration call to add our registry entries to the SolidWorks add-in registry
        /// </summary>
        /// <param name="t"></param>

        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format($"SOFTWARE/SolidWorks/AddIns/{0:b}", t.GUID);

            // Create our registry folder for the add-in
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                // Load add-in when SolidWorks opens
                rk.SetValue(null, 1);

                // Set Solidworks add-in title and description 
                rk.SetValue("Title", "My SwAddin");
                rk.SetValue("Description", "All your pixels are belong to us!");
            }
        }

        /// <summary>
        /// The COM unregister call to remove our custom entries we added in the COM register function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE/sOLIDwORKS/aDDiNS/{0:B}", t.GUID);

            // Remove our registry entry
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }
        #endregion
    }
}
// a small change 39:15 
