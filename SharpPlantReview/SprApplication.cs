//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Reflection;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides methods and properties for interacting with SmartPlant Review.
    /// </summary>
    public class SprApplication : IDisposable
    {
        public sealed class SprApplicationWindows
        {
            /// <summary>
            ///     The parent Application reference.
            /// </summary>
            public SprApplication Application { get; private set; }

            /// <summary>
            ///     The primary SmartPlant Review application window.
            /// </summary>
            public SprWindow ApplicationWindow
            {
                get
                {
                    if (applicationWindow == null)
                        applicationWindow = GetWindow(SprWindowType.ApplicationWindow);
                    return applicationWindow;
                }
                set
                {
                    applicationWindow = value;
                }
            }
            private SprWindow applicationWindow;

            /// <summary>
            ///     The Elevation window inside the SmartPlant Review application.
            /// </summary>
            public SprWindow ElevationWindow
            {
                get
                {
                    if (elevationWindow == null)
                        elevationWindow = GetWindow(SprWindowType.ElevationWindow);
                    return elevationWindow;
                }
                set
                {
                    applicationWindow = value;
                }
            }
            private SprWindow elevationWindow;

            /// <summary>
            ///     The Main view window inside the SmartPlant Review application.
            /// </summary>
            public SprWindow MainWindow
            {
                get
                {
                    if (mainWindow == null)
                        mainWindow = GetWindow(SprWindowType.MainWindow);
                    return mainWindow;
                }
                set
                {
                    mainWindow = value;
                }
            }
            private SprWindow mainWindow;

            /// <summary>
            ///     The Plan view window inside the SmartPlant Review application.
            /// </summary>
            public SprWindow PlanWindow
            {
                get
                {
                    if (planWindow == null)
                        planWindow = GetWindow(SprWindowType.PlanWindow);
                    return planWindow;
                }
                set
                {
                    planWindow = value;
                }
            }
            private SprWindow planWindow;

            /// <summary>
            ///     The Text window inside the SmartPlant Review application.
            /// </summary>
            public SprTextWindow TextWindow
            {
                get
                {
                    if (textWindow == null)
                        textWindow = (SprTextWindow)GetWindow(SprWindowType.TextWindow);
                    return textWindow;
                }
                set
                {
                    textWindow = value;
                }
            }
            private SprTextWindow textWindow;

            internal SprApplicationWindows(SprApplication application)
            {
                Application = application;

                applicationWindow = GetWindow(SprWindowType.ApplicationWindow);
                elevationWindow = GetWindow(SprWindowType.ElevationWindow);
                mainWindow = GetWindow(SprWindowType.MainWindow);
                planWindow = GetWindow(SprWindowType.PlanWindow);
                textWindow = (SprTextWindow)GetWindow(SprWindowType.TextWindow);
            }

            private SprWindow GetWindow(SprWindowType type)
            {
                if (!Application.IsConnected)
                    throw SprExceptions.SprNotConnected;

                if (type == SprWindowType.TextWindow)
                    return new SprTextWindow(Application);
                return new SprWindow(Application, type);
            }
        }

        #region Properties

        private bool disposed;

        /// <summary>
        ///     The static Application used to set the class parent object.
        /// </summary>
        internal static SprApplication ActiveApplication;

        /// <summary>
        ///     The collection of active running SprProcesses.
        /// </summary>
        public static Process[] SprProcesses
        {
            get { return Process.GetProcessesByName("spr"); }
        }

        /// <summary>
        ///     The active COM reference to the DrAPi class.
        /// </summary>
        internal dynamic DrApi;

        /// <summary>
        ///     The active session database.
        /// </summary>
        public DataSet MdbDatabase
        {
            get { return mdbDatabase ?? (mdbDatabase = GetMdbDatabase()); }
            internal set { mdbDatabase = value; }
        }
        private DataSet mdbDatabase;

        /// <summary>
        ///     The SprAnnotation collection from the current MDB database.
        /// </summary>
        public SprAnnotationCollection Annotations
        {
            get { return annotations ?? (annotations = new SprAnnotationCollection(this)); }
            set { annotations = value; }
        }
        private SprAnnotationCollection annotations;

        /// <summary>
        ///     The default properties used when an SprSnapshot is omitted from snapshot methods.
        /// </summary>
        public SprSnapShot DefaultSnapshot { get; set; }

        /// <summary>
        ///     Gets a list of the design files loaded into the active review session.
        /// </summary>
        public List<string> DesignFiles
        {
            get { return designFiles ?? (designFiles = GetDesignFiles()); }
        }
        private List<string> designFiles;

        /// <summary>
        ///     Determines if the SmartPlant Review application is busy.
        /// </summary>
        public bool IsBusy
        {
            get { return CheckIsBusy(); }
        }

        /// <summary>
        ///     Determines if an active connection to SmartPlant Review is established.
        /// </summary>
        public bool IsConnected
        {
            // Check if the DrApi is set
            get { return DrApi != null; }
        }

        /// <summary>
        ///     The last error message returned from the most recent DrApi function call.
        /// </summary>
        public string LastError { get; internal set; }

        /// <summary>
        ///     Gets the MDB path to the active review session.
        /// </summary>
        public string MdbPath
        {
            get { return mdbPath ?? (mdbPath = GetMdbPath()); }
        }
        private string mdbPath;

        /// <summary>
        ///     Gets the next annotation number to be used.
        /// </summary>
        public int NextAnnotation
        {
            get
            {
                if (nextAnnotation == 0)
                    nextTag = GetNextAnnotation();
                return nextAnnotation;
            }
            set 
            {
                SetNextAnnotation(value);
                nextAnnotation = value;
            }
        }
        private int nextAnnotation;

        /// <summary>
        ///     Gets the next tag number to be used.
        /// </summary>
        public int NextTag
        {
            get
            {
                if (nextTag == 0)
                    nextTag = GetNextTag();
                return nextTag;
            }
            set
            {
                SetNextTag(value);
                nextTag = value;
            }
        }
        private int nextTag;

        /// <summary>
        ///     Gets the process ID of the SmartPlant Review application.
        /// </summary>
        public IntPtr ProcessId
        {
            get
            {
                if (processId == IntPtr.Zero)
                    processId = GetProcessId();
                return processId;
            }
        }
        private IntPtr processId;

        /// <summary>
        ///     Gets the filename of the active review session.
        /// </summary>
        public string SessionName
        {
            get { return sessionName ?? (sessionName = GetSessionName()); }
        }
        private string sessionName;

        /// <summary>
        ///     The returned result from the most recent DrApi function call.
        /// </summary>
        public int SprStatus
        {
            get { return sprStatus; }
            internal set
            {
                sprStatus = value;

                // Handle the errors
                SprUtilities.ErrorCheck(value);
            }
        }
        private int sprStatus;

        /// <summary>
        ///     The SprTag collection from the current MDB database.
        /// </summary>
        public SprTagCollection Tags
        {
            get { return tags ?? (tags = new SprTagCollection(this)); }
            set { tags = value; }
        }
        private SprTagCollection tags;

        /// <summary>
        ///     Gets the version of the running instance of SmartPlant Review.
        /// </summary>
        public string Version
        {
            get { return version; }
        }
        private readonly string version;

        /// <summary>
        ///     Represents a structure comntaining the SprApplication Windows.
        /// </summary>
        public SprApplicationWindows Windows { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        ///     Creates a new SprApplication class.  Will automatically connect to a
        ///     single SmartPlant Review process if available.
        /// </summary>
        public SprApplication()
        {
            // Set the static application for class parent referencing
            ActiveApplication = this;

            // If only one instance of Spr is running
            if (SprProcesses.Length == 1)

                // Connect to the SPR instance automatically
                Connect();

            // Create the default snapshot format
            DefaultSnapshot = new SprSnapShot
            {
                AntiAlias = 3,
                OutputFormat = SprSnapshotFormat.Jpg,
                AspectOn = true,
                Scale = 1
            };

            // Set the default snapshot directories
            SprSnapShot.TempDirectory = Environment.GetEnvironmentVariable("TEMP");
            SprSnapShot.DefaultDirectory = Environment.SpecialFolder.MyPictures.ToString();

            // Set the startup fields
            version = GetVersion();
            mdbPath = GetMdbPath();
            mdbDatabase = GetMdbDatabase();
            sessionName = GetSessionName();
            nextTag = GetNextTag();
            nextAnnotation = GetNextAnnotation();
            processId = GetProcessId();
            sprStatus = 0;
            designFiles = GetDesignFiles();
            tags = new SprTagCollection(this);
            
            // Windows
            Windows = new SprApplicationWindows(this);
        }

        /// <summary>
        ///     SprApplication deconstructor/finalizer.
        /// </summary>
        ~SprApplication()
        {
            Dispose(false);
        }

        #endregion

        #region Private Methods

        private string GetSessionName()
        {
            if (!IsConnected)
                return null;

            string sessionName = null;

            // Set the DrApi global options to return the file name
            //SprStatus = DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);
            GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);
            
            // Get the session file name
            //SprStatus = DrApi.FileNameFromNumber(0, ref sessionName);
            var result = Run(SprNativeMethods.FileNameFromNumber, 0, sessionName);
            sessionName = (string)result[1];

            return sessionName;
        }
        private string GetMdbPath()
        {
            if (!IsConnected)
                return null;

            // params for holding the return data
            var dirPath = string.Empty;
            var mdbName = string.Empty;

            // Set the DrApi global options to return the file name
            //SprStatus = DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);
            GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);

            // MDB Name
            //SprStatus = DrApi.FileNameFromNumber(1, ref mdbName);
            var result = Run(SprNativeMethods.FileNameFromNumber, 1, mdbName);
            mdbName = (string)result[1];
            
            // MDB Directory
            //SprStatus = DrApi.FilePathFromNumber(1, ref dirPath);
            result = Run(SprNativeMethods.FilePathFromNumber, 1, dirPath);
            dirPath = (string)result[1];

            if (dirPath != null && mdbName != null)
                return Path.Combine(dirPath, mdbName);

            return null;
        }
        private IntPtr GetProcessId()
        {
            if (!IsConnected)
                return IntPtr.Zero;

            var procId = default(uint);
            //SprStatus = DrApi.ProcessIdGet(out procId);
            var result = Run(SprNativeMethods.ProcessIdGet, procId);

            return (IntPtr)result[0];
        }
        private bool CheckIsBusy()
        {
            if (!IsConnected)
                return false;

            var sw = new Stopwatch();
            sw.Start();

            // Wait until the application is idle (10 ms max)
            Process.GetProcessById((int)ProcessId).WaitForInputIdle(10);
            sw.Stop();

            // Return true if the application waited the full test period
            return (sw.ElapsedMilliseconds >= 10);
        }
        private int GetNextAnnotation()
        {
            // Get the next annotation number
            var siteTable = MdbDatabase.Tables["site_table"];
            return (int)siteTable.Rows[0]["next_text_anno_id"];
        }
        private int GetNextTag()
        {
            // Get the next tag number
            var returnTag = default(uint);
            //SprStatus = DrApi.TagNextNumber(out returnTag, 0);
            var result = Run(SprNativeMethods.TagNextNumber, returnTag);

            return Convert.ToInt32(result[0]);
        }
        private string GetVersion()
        {
            if (!IsConnected)
                return null;

            string vers = null;

            // Get the version of the SPR Application
            //SprStatus = DrApi.Version(ref vers);
            var result = Run(SprNativeMethods.Version, vers);

            return (string)result[0];
        }
        private DataSet GetMdbDatabase()
        {
            // TODO:  Add a DbMethod for returning a set of tables so that a single Db connection can be used 
            var tables = new string[] { "tag_data", "site_table", "text_annotations", "text_annotation_types" };

            var returnSet = new DataSet("MDB Database");

            for (int i = 0; i < tables.Length; i++)
                returnSet.Tables.Add(DbMethods.GetDbTable(MdbPath, tables[i]));

            return returnSet;
        }
        private List<string> GetDesignFiles()
        {
            GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

            var fileCount = default(uint);
            //SprStatus = DrApi.FileCountGet(out fileCount);
            var result = Run(SprNativeMethods.FileCountGet, fileCount);
            fileCount = Convert.ToUInt32(result[0]);

            var returnList = new List<string>();

            for (int i = 0; i < fileCount; i++)
            {
                string curName;
                string curPath;
                SprStatus = DrApi.FileNameFromNumber(i, out curName);
                SprStatus = DrApi.FilePathFromNumber(i, out curPath);

                returnList.Add(Path.Combine(curPath, curName));
            }

            return returnList;
        }

        private void SetNextAnnotation(int annoId)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            var siteTable = MdbDatabase.Tables["site_table"];
            var tagTable = MdbDatabase.Tables["text_annotations"];
            if (tagTable.Rows.Count > 0)

                // Set the next tag to the highest tag value + 1
                siteTable.Rows[0]["next_anno_id"] =
                    Convert.ToInt32(tagTable.Rows[tagTable.Rows.Count - 1]["tag_unique_id"]) + 1;
            else

                // Set the next tag to 1
                siteTable.Rows[0]["next_anno_id"] = 1;

            DbMethods.UpdateDbTable(MdbPath, siteTable);
        }
        private void SetNextTag(int tagId)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            var siteTable = MdbDatabase.Tables["site_table"];
            var tagTable = MdbDatabase.Tables["tag_data"];
            if (tagTable.Rows.Count > 0)

                // Set the next tag to the highest tag value + 1
                siteTable.Rows[0]["next_tag_id"] =
                    Convert.ToInt32(tagTable.Rows[tagTable.Rows.Count - 1]["tag_unique_id"]) + 1;
            else

                // Set the next tag to 1
                siteTable.Rows[0]["next_tag_id"] = 1;

            DbMethods.UpdateDbTable(MdbPath, siteTable);
        }

        /// <summary>
        ///     Runs the specified DrApi method through the active AprApplication.
        /// </summary>
        /// <param name="method">The SprNativeMethod to invoke.</param>
        /// <param name="args">The arguments passed to the method.</param>
        /// <returns>The object array returned from the method.</returns>
        public object[] Run(SprNativeMethods method, params object[] args)
        {
            SprStatus =  DrApi.InvokeMember(method.ToString(), BindingFlags.InvokeMethod | BindingFlags.Public,
                                           null, null, args);

            return args;
        }

        #endregion

        #region Public Methods

        #region General

        /// <summary>
        ///     Brings the SmartPlant Review application to the foreground.
        /// </summary>
        public void Activate()
        {
            NativeWin32.SetForegroundWindow(Windows.ApplicationWindow.hWnd);
        }

        /// <summary>
        ///     Connects to a running instance of SmartPlant Review.
        /// </summary>
        /// <returns>Boolean indicating success or failure of the operation.</returns>
        public bool Connect()
        {
            DrApi = Activator.CreateInstance(SprImportedTypes.DrApi);
            return IsConnected;
        }

        /// <summary>
        ///     Releases the connection to the SprApplication.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    if (DrApi != null)
                        Marshal.ReleaseComObject(DrApi);
                }

                disposed = true;
                DrApi = null;
            }
        }

        /// <summary>
        ///     Attempts to open the specified session in the fileName.
        ///     Any active session will first be closed.
        /// </summary>
        /// <param name="fileName">The full path of the session to load.</param>
        public void Open(string fileName)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Try opening the file using the Api method call
            SprStatus = DrApi.SessionAttach(fileName);

            // Set the session vars
            mdbPath = GetMdbPath();
            mdbDatabase = GetMdbDatabase();
            sessionName = GetSessionName();
            nextTag = GetNextTag();
            nextAnnotation = GetNextAnnotation();
            designFiles = GetDesignFiles();
            tags = new SprTagCollection(this);
        }

        /// <summary>
        ///     Exports the current session to a VUE file format.
        ///     Compatible only with SPR versions 9 and above.
        /// </summary>
        /// <param name="vueName">The full path of the vue file to be exported.</param>
        public void Export(string vueName)
        {
            // Check version compatibility
            int vers = int.Parse(Version.Substring(0, 2));
            if (vers < 9)
                throw SprExceptions.SprVersionIncompatibility;

            // Export the current session to VUE
            SprStatus = DrApi.ExportVue(vueName, 0);
        }

        /// <summary>
        ///     Exits the SmartPlant Review application.
        /// </summary>
        public void Exit()
        {
            if (!IsConnected)
                return;

            // Exit the SPR application
            SprStatus = DrApi.ExitViewer();
        }

        /// <summary>
        ///     Refreshes the session data as if selected from the application menu.
        /// </summary>
        public void RefreshData()
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            Activate();

            // Send the update command
            SendKeys.SendWait("%TR");
        }

        /// <summary>
        ///     Sets the value of a specified global option in SmartPlant Review.
        /// </summary>
        /// <param name="option">The integer of the global option to set.</param>
        /// <param name="value">The integer value to set the global option to.</param>
        public void GlobalOptionsSet(int option, int value)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Set the global option to the provided value
            SprStatus = DrApi.GlobalOptionsSet(option, value);
        }

        /// <summary>
        ///     Gets the value of a specified global option in SmartPlant Review.
        /// </summary>
        /// <param name="option">The integer value of the option to retrieve.</param>
        /// <returns>Integer value representing the current state of the option.</returns>
        public double GlobalOptionsGet(int option)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            double returnVal;

            // Get the global option value
            SprStatus = DrApi.GlobalOptionsGet(option, out returnVal);

            return returnVal;
        }

        /// <summary>
        ///     Highlights a specified object in the main SmartPlant Review application Window
        /// </summary>
        /// <param name="objectId">Object Id of the of the entity to be highlighted.</param>
        public void HighlightObject(int objectId)
        {
            // Highlight the object in SPR
            SprStatus = DrApi.HighlightObject(objectId, 1, 0);
        }

        /// <summary>
        ///     Clears all highlighting from the main SmartPlant Review application window.
        /// </summary>
        public void HighlightClear()
        {
            // Clear the highlighting in SPR 
            SprStatus = DrApi.HighlightExit(1);

            // Refresh the main window
            SprStatus = DrApi.ViewUpdate(1);
        }

        #endregion

        #region Point Select

        /// <summary>
        ///     Prompts a user to select a point inside SmartPlant Review.
        ///     Retrieves the physical point selected on screen.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <returns>SprPoint3D representing the selected point.</returns>
        public SprPoint3D GetPoint(string prompt)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            var returnPoint = new SprPoint3D();
            int abortFlag;

            Activate();

            // Prompt the user for a 3D point inside SPR
            SprStatus = DrApi.PointLocateDbl(prompt, out abortFlag, ref returnPoint.DrPointDbl);

            // Return null if the locate operation was aborted
            return abortFlag != 0 ? returnPoint : null;
        }

        /// <summary>
        ///     Prompts a user to select a point inside SmartPlant Review.
        ///     Retrieves the physical point selected on screen.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <param name="targetPoint">The point used to calculate depth.</param>
        /// <returns></returns>
        public SprPoint3D GetPoint(string prompt, SprPoint3D targetPoint)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            var returnPoint = new SprPoint3D();
            int abort;
            int objId;
            const int flag = 0;

            Activate();

            // Prompt the user for a 3D point inside SPR
            SprStatus = DrApi.PointLocateExtendedDbl(prompt, out abort, ref returnPoint.DrPointDbl,
                                                         ref targetPoint.DrPointDbl, out objId, flag);

            // Return null if the locate operation was aborted
            return abort != 0 ? returnPoint : null;
        }

        /// <summary>
        ///     Prompts a user to select an object inside SmartPlant Review.
        ///     Retrieves the Object Id of the selected object.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <returns>Integer value representing the selected Object Id.</returns>
        public int GetObjectId(string prompt)
        {
            var pt = new SprPoint3D();
            return GetObjectId(prompt, ref pt);
        }

        /// <summary>
        ///     Prompts a user to select an object inside SmartPlant Review.
        ///     Retrieves the Object Id of the selected object.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <param name="refPoint">The returned reference point representing the selected point.</param>
        /// <returns>Integer value representing the selected Object Id.</returns>
        public int GetObjectId(string prompt, ref SprPoint3D refPoint)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            const int filterFlag = 0;
            var returnId = -1;

            Activate();

            // Prompt the user for a 3D point inside SPR
            SprStatus = DrApi.ObjectLocateDbl(prompt, filterFlag, out returnId, ref refPoint.DrPointDbl);

            return returnId;
        }

        /// <summary>
        ///     Uses an Object Id to retrieve detailed object information.
        /// </summary>
        /// <param name="objectId">The ObjectId that is looked up inside SmartPlant Review.</param>
        /// <returns>SprObjectData object containing the retrieved information.</returns>
        public SprObject GetObjectData(int objectId)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            if (objectId == 0)
                return null;

            var returnData = new SprObject();
            returnData.Id = objectId;

            // Get the DataDbl object
            SprStatus = DrApi.ObjectDataGetDbl(objectId, 2, ref returnData.DrObjectDataDbl);

            // Iterate through the labels
            string lblName = string.Empty, lblValue = string.Empty;
            for (var i = 0; i < returnData.DrObjectDataDbl.LabelDataCount; i++)
            {
                SprStatus = DrApi.ObjectDataLabelGet(ref lblName, ref lblValue, i);
                if (!returnData.Labels.ContainsKey(lblName))
                    returnData.Labels.Add(lblName, lblValue);
            }
            return returnData;
        }

        /// <summary>
        ///     Prompts a user to select an object inside SmartPlant Review.
        ///     Retrieves object information from the selected object.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <returns>The SprObjectData object containing the retrieved information.</returns>
        public SprObject GetObjectData(string prompt)
        {
            return GetObjectData(prompt, false);
        }

        /// <summary>
        ///     Prompts a user to select an object inside SmartPlant Review.
        ///     Retrieves object information from the selected object.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <param name="singleObjects">Indicates if SmartPlant Review locates grouped objects individually.</param>
        /// <returns>The SprObjectData object containing the retrieved information.</returns>
        public SprObject GetObjectData(string prompt, bool singleObjects)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Get the ObjectID on screen
            var objId = GetObjectId(prompt);

            return GetObjectData(objId);
        }

        /// <summary>
        ///     Searches through the active session for objects matching the supplied search criteria.
        /// </summary>
        /// <param name="criteria">The search criteria used to find matching objects.</param>
        /// <returns>A list of ObjectIds representing the matching objects.</returns>
        public List<int> ObjectDataSearch(string criteria)
        {
            int itemCount;
            SprStatus = DrApi.ObjectDataSearch(criteria, 0, out itemCount);

            var returnIds = new List<int>();
            for (int i = 0; i < itemCount; i++)
            {
                int curId;
                SprStatus = DrApi.ObjectDataSearchIdGet(out curId, i);
                if (curId != 0)
                    returnIds.Add(curId);

            }

            return returnIds;
        }
        #endregion

        #region Annotation

        /// <summary>
        ///     Toggles annotation display in the SmartPlant Review application main window.
        /// </summary>
        /// <param name="visible">Determines the annotation visibility state.</param>
        //public void Annotations_Display(bool visible)
        //{
        //    // Throw an expection if not connected
        //    if (!IsConnected) throw SprExceptions.SprNotConnected;

        //    // Create the params
        //    var visValue = Convert.ToInt32(visible);

        //    // Create the view object
        //    dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

        //    // Throw an exception if the DrViewDbl is null
        //    if (objViewdataDbl == null) throw SprExceptions.SprObjectCreateFail;

        //    // Set the view object as the SPR Application main view
        //    SprStatus = DrApi.ViewGetDbl(0, ref objViewdataDbl);

        //    // Apply the updated annotation display
        //    objViewdataDbl.AllAnnotationsDisplay = visValue;

        //    // Update the global annotation visibility properties
        //    SprStatus = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDisplay, visValue);
        //    SprStatus = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoTextDisplay, visValue);
        //    SprStatus = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDataDisplay, visValue);

        //    // Update the main view in SPR
        //    SprStatus = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        //}

        /// <summary>
        ///     Creates a new data field in the Mdb text_annotations table.
        ///     Returns true if the field already exists.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <returns>Indicates the success or failure of the table modification.</returns>
        //public bool Annotations_AddDataField(string fieldName)
        //{
        //    // Throw an exception if not connected
        //    if (!IsConnected) throw SprExceptions.SprNotConnected;

        //    // Add the tag field to the MDB database
        //    return DbMethods.AddDbField(MdbPath, "text_annotations", fieldName);
        //}

        /// <summary>
        ///     Prompts a user to place a new annotation in the SmartPlant Review main view.
        /// </summary>
        /// <param name="anno">Reference SprAnnotation containing the annotation information.</param>
        public void Annotations_Place(ref SprAnnotation anno)
        {
            // Throw exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the leader point
            var leaderPoint = new SprPoint3D();

            // Get the annotation leader point
            int objId = GetObjectId("SELECT A POINT ON AN OBJECT TO LOCATE THE ANNOTATION", ref leaderPoint);

            // Exit if the object selection failed/canceled
            if (objId == 0)
            {
                Windows.TextWindow.Text = "Annotation placement canceled.";
                return;
            }

            // Get the annotation center point using the leaderpoint for depth calculation
            var centerPoint = GetPoint("SELECT THE CENTER POINT FOR THE ANNOTATION LABEL", leaderPoint);

            // Exit if the centerpoint selection failed/canceled
            if (centerPoint == null)
            {
                Windows.TextWindow.Text = "Annotation placement canceled.";
                return;
            }

            // Create a reference to the DrAnnotation
            var drAnno = anno.DrAnnotationDbl;

            // Set the annotation points
            drAnno.LeaderPoint = leaderPoint.DrPointDbl;
            drAnno.CenterPoint = centerPoint.DrPointDbl;

            // Place the annotation on screen
            int annoId;
            SprStatus = DrApi.AnnotationCreateDbl(anno.Type, ref drAnno, out annoId);

            // Link the located object to the annotation
            SprStatus = DrApi.AnnotationDataSet(annoId, anno.Type, ref drAnno, ref objId);

            // Retrieve the placed annotation data
            anno = Annotations_Get(anno.Id);

            // Add an ObjectId field
            //Annotations_AddDataField("object_id");

            // Save the ObjectId to the annotation data
            anno.Data["object_id"] = objId;

            // Update the annotation
            Annotations_Update(anno);

            // Update the text window
            Windows.TextWindow.Title = string.Format("Annotation {0}", anno.Id);
            Windows.TextWindow.Text = anno.Text;

            // Update the main view
            SprStatus = DrApi.ViewUpdate(1);
        }

        public void Annotations_EditLeader(int annoNo)
        {

        }
        public void Annotations_EditLeader(ref SprAnnotation annotation)
        {

        }

        /// <summary>
        ///     Adds an annotation directly into the Mdb annotation table.
        /// </summary>
        /// <param name="anno">SprAnnotation containing the annotation information.</param>
        public void Annotations_Add(SprAnnotation anno)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Adds a list of annotations into the Mdb annotation table.
        /// </summary>
        /// <param name="annotations">List of SprAnnotation to be added to the annotation table.</param>
        public void Annotations_Add(List<SprAnnotation> annotations)
        {
            // Add each annotation
            foreach (var anno in annotations)
                Annotations_Add(anno);
        }

        /// <summary>
        ///     Prompts a user to select an annotation in the SmartPlant Review main view.
        /// </summary>
        /// <param name="type">The string type of the annotation to be selected.</param>
        /// <returns>SprAnnotation containing the selected annotation information.</returns>
        public SprAnnotation Annotations_Select(string type)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the params
            int annoId;

            // Create the return annotation object
            var anno = new SprAnnotation();

            // Set the SPR application visible
            Activate();

            // Prompt the user to select the annotation
            var msg = string.Format("SELECT THE DESIRED {0} ANNOTATION", type.ToUpper());
            SprStatus = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Return null if the annotation locate failed
            if (annoId == 0) return null;

            // Set the annotation ID
            //anno.Id = annoId;

            // Get the associated object ID
            int assocId;
            var drAnno = anno.DrAnnotationDbl;
            SprStatus = DrApi.AnnotationDataGet(annoId, type, ref drAnno, out assocId);

            // Return null if the associated object id is zero
            if (assocId == 0) return null;

            // Set the assiciated object
            //anno.AssociatedObject = GetObjectData(assocId);

            // Return the completed annotation
            return anno;
        }

        /// <summary>
        ///     Prompts a user to delete an annotation in the SmartPlant Review main view.
        /// </summary>
        /// <param name="type">The string type of the annotation to delete.</param>
        public void Annotations_Delete(string type)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the params
            int annoId;

            // Set the SPR application visible
            Activate();

            // Prompt the user to select the annotation
            var msg = string.Format("SELECT THE {0} ANNOTATION TO DELETE", type.ToUpper());
            SprStatus = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Return if the annotation locate was unsuccessful
            if (annoId == 0) return;

            // Delete the selected annotation
            SprStatus = DrApi.AnnotationDelete(type, annoId, 0);

            // Update the main view
            SprStatus = DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Deletes all the annotations of a specified type from the active SmartPlant Review session.
        /// </summary>
        /// <param name="type">The string type of the annotations to delete.</param>
        public void Annotations_DeleteType(string type)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Delete all annotations matching the provided type
            SprStatus = DrApi.AnnotationDeleteAll(type, 0);

            // Update the main view
            SprStatus = DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Deletes all the annotations from the active SmartPlant Review session.
        /// </summary>
        public void Annotations_DeleteAll()
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Get the annotation types
            var typeTable = DbMethods.GetDbTable(MdbPath, "text_annotation_types");

            // Exit if the type table is null
            if (typeTable == null) return;

            // If no types exist
            if (typeTable.Rows.Count == 0) return;

            // Iterate through each annotation type
            for (var i = typeTable.Rows.Count - 1; i >= 0; i--)
            {
                // Delete all annotations matching the current type
                Annotations_DeleteType(typeTable.Rows[i]["name"].ToString());
            }

            // Update the main view
            SprStatus = DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Retrieves the desired annotation from the Mdb text_annotations table.
        /// </summary>
        /// <param name="annoId">Unique Id of the annotation to retrieve.</param>
        /// <returns>SprAnnotation containing the retirned annotation information.</returns>
        public SprAnnotation Annotations_Get(int annoId)
        {
            // Throw an exception if not connected
            //if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the new tag
            var anno = new SprAnnotation();

            // Retrieve the site table
            var annoTable = DbMethods.GetDbTable(MdbPath, "text_annotations");

            // Return null if the table retrieval failed
            if (annoTable == null) return null;

            // Return null if no annotations exist
            if (annoTable.Rows.Count == 0) return null;

            // Create the row filter for the desired tag
            var rowFilter = annoTable.Select(string.Format("id = '{0}'", annoId));

            // Throw an exception if the annotation was not found
            if (rowFilter.Length == 0) throw SprExceptions.SprAnnotationNotFound;

            // Iterate through each column
            foreach (DataColumn col in annoTable.Columns)
            {
                // Add the key/value from the first filtered row to the dictionary
                anno.Data[col.ColumnName] = rowFilter[0][col];
            }

            // Set the tag as placed
            anno.IsPlaced = true;

            // Return the tag
            return anno;
        }

        /// <summary>
        ///     Updates annotation data directly in the Mdb text_annotations table.
        /// </summary>
        /// <param name="anno">SprAnnotation containing the annotation information.</param>
        /// <returns>Indicates the success or failure of the text_annotations modification.</returns>
        public bool Annotations_Update(SprAnnotation anno)
        {
            // Throw an exception if not connected
            //if (!IsConnected)throw SprExceptions.SprNotConnected;

            // Retrieve the site table
            var annoTable = DbMethods.GetDbTable(MdbPath, "text_annotations");

            // Return false if the table is null
            if (annoTable == null) return false;

            // Return false if no tags exist
            if (annoTable.Rows.Count == 0) return false;

            // Create the row filter for the specified tag
            var rowFilter = string.Format("id = {0}", anno.Id);
            var tblFilter = annoTable.Select(rowFilter);

            // Iterate through each dictionary key/value pair
            foreach (var kvp in anno.Data)

                // Set the values for the selected tag
                tblFilter[0][kvp.Key] = kvp.Value;

            // Return the result of the table update
            return DbMethods.UpdateDbTable(MdbPath, rowFilter, annoTable);
        }

        /// <summary>
        ///     Updates annotation information directly in the Mdb text_annotations table, queueable from the threadpool.
        /// </summary>
        /// <param name="stateInfo">SprAnnotation passed as an object per WaitCallback requirements.</param>
        public void Annotations_Update(object stateInfo)
        {
            // Cast the threading object 
            var anno = stateInfo as SprAnnotation;

            // Return if the annotation is null
            if (anno == null) return;

            // Retrieve the annotation table
            var annoTable = DbMethods.GetDbTable(MdbPath, "text_annotations");

            // Return if the table is null
            if (annoTable == null) return;

            // Return if no annotations exist
            if (annoTable.Rows.Count == 0) return;

            // Create the row filter for the specified tag
            var rowFilter = string.Format("id = {0}", anno.Id);
            var tblFilter = annoTable.Select(rowFilter);

            // Iterate through each dictionary key/value pair
            foreach (var kvp in anno.Data)

                // Set the values for the selected annotation
                tblFilter[0][kvp.Key] = kvp.Value;

            // Push the the updated table
            DbMethods.UpdateDbTable(MdbPath, rowFilter, annoTable);
        }

        #endregion

        #region Snapshot

        /// <summary>
        ///     Captures the active SmartPlant Reviews session main view and 
        ///     writes it an image of the given filename and format.
        /// </summary>
        /// <param name="snapShot">SprSnapshot containing the snapshot settings.</param>
        /// <param name="imageName">Name of the final output image.</param>
        /// <param name="outputDir">Directory to save the snapshot to.</param>
        /// <returns></returns>
        public Image TakeSnapshot(string imageName, string outputDir, SprSnapShot snapShot = null)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Get the default snapshot if none was supplied
            var snap = snapShot ?? DefaultSnapshot;           

            // Get the current backface/endcap settings
            var orgBackfaces = GlobalOptionsGet(SprConstants.SprGlobalBackfacesDisplay);
            var orgEndcaps = GlobalOptionsGet(SprConstants.SprGlobalEndcapsDisplay);

            // Turn on view backfaces/endcaps as needed
            if (orgBackfaces == 0)
                GlobalOptionsSet(SprConstants.SprGlobalBackfacesDisplay, 1);
            if (orgEndcaps == 0)
                GlobalOptionsSet(SprConstants.SprGlobalEndcapsDisplay, 1);

            // (.BMP is forced before conversions)
            var imgPath = Path.Combine(outputDir, string.Format("{0}.bmp", imageName));

            // Take the snapshot
            SprStatus = DrApi.SnapShot(imgPath, snap.Flags, snap.DrSnapShot, 0);

            // Wait until finished
            while (IsBusy)
                Thread.Sleep(100);

            // Reset the original settings if applicable
            if (orgBackfaces == 0)
                GlobalOptionsSet(SprConstants.SprGlobalBackfacesDisplay, 1);
            if (orgEndcaps == 0)
                GlobalOptionsSet(SprConstants.SprGlobalEndcapsDisplay, 1);

            // Return false if the snapshot doesn't exist
            if (!File.Exists(imgPath))
                return null;

            // Format the snapshot if required
            if (snapShot.OutputFormat != SprSnapshotFormat.Bmp)
                SprSnapShot.FormatSnapshot(imgPath, snap.OutputFormat);

            return Image.FromFile(imgPath);
        }

        /// <summary>
        ///     Compatible only with SPR versions 9 and above.
        /// </summary>
        /// <param name="quality"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public void ExportPDF(int quality, string path)
        {
            // Check version compatibility
            int vers = int.Parse(Version.Substring(0, 2));
            if (vers < 9)
                throw SprExceptions.SprVersionIncompatibility;

            SprStatus = DrApi.ExportPDF(path, quality, 1, 1, 1);
        }

        #endregion

        #region Views

        public void SetCenterPoint(double east, double north, double elevation)
        {
            SetCenterPoint(new SprPoint3D(east, north, elevation));
        }

        public SprPoint3D GetCenterPoint()
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Create the DrViewDbl
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            // Set the view object as the SPR Application main view
            SprStatus = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Return the centerpoint
            return new SprPoint3D(objViewdataDbl.CenterUorPoint);
        }

        public void SetCenterPoint(SprPoint3D centerPoint)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Create the DrViewDbl
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            // Set the view object as the SPR Application main view
            SprStatus = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated centerpoint
            objViewdataDbl.CenterUorPoint = centerPoint.DrPointDbl;

            // Update the main view in SPR
            SprStatus = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        public void SetEyePoint(double east, double north, double elevation)
        {
            SetEyePoint(new SprPoint3D(east, north, elevation));
        }

        public void SetEyePoint(SprPoint3D eyePoint)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Create the DrViewDbl
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            // Set the view object as the SPR Application main view
            SprStatus = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated eyepoint
            objViewdataDbl.EyeUorPoint = eyePoint.DrPointDbl;

            // Update the main view in SPR
            SprStatus = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        #endregion

        #endregion
    }
}