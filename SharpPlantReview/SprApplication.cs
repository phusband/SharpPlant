//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides methods and properties for interacting with SmartPlant Review.
    /// </summary>
    public class SprApplication : IDisposable
    {
        #region Application Properties

        /// <summary>
        ///     The static Application used to set the class parent object.
        /// </summary>
        internal static SprApplication ActiveApplication;

        private bool _disposed;

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
        ///     Determines if an active connection to SmartPlant Review is established.
        /// </summary>
        public bool IsConnected
        {
            // Check if the DrApi is set
            get { return DrApi != null; }
        }

        /// <summary>
        ///     Gets the MDB path to the active review session.
        /// </summary>
        public string MdbPath
        {
            get
            {
                // Return null if not connected
                if (!IsConnected) return null;

                // params for holding the return data
                var dirPath = string.Empty;
                var mdbName = string.Empty;

                // Set the global option to retrieve database file info
                DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);

                // Get the directory path
                DrApi.FilePathFromNumber(1, ref dirPath);

                // Get the MDB FileName
                DrApi.FileNameFromNumber(1, ref mdbName);

                // Reset the global variable
                DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

                // Check if the values are set
                if (dirPath != null && mdbName != null)
                {
                    // Return the full MDB path
                    return Path.Combine(dirPath, mdbName);
                }

                return null;
            }
        }

        /// <summary>
        ///     Gets the filename of the active review session.
        /// </summary>
        public string SessionName
        {
            get
            {
                // Resturn null if not connected
                if (!IsConnected) return null;

                // Create the return parameter
                string vueName = null;

                // Set the global variable
                DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);

                // Get the VUE file name
                DrApi.FileNameFromNumber(0, ref vueName);

                // Reset the global variable
                DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

                // Return the path
                return vueName;
            }
        }

        /// <summary>
        ///     Gets a list of the design files loaded into the active review session.
        /// </summary>
        public List<string> DesignFiles
        {
            get
            {  
                int fileCount;
                int sprResult = DrApi.FileCountGet(out fileCount);

                GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

                var returnList = new List<string>();

                for (int i = 0; i < fileCount; i++)
                {
                    string curName;
                    string curPath;
                    sprResult = DrApi.FileNameFromNumber(i, out curName);

                    sprResult = DrApi.FilePathFromNumber(i, out curPath);

                    returnList.Add(string.Format("{0}{1}", curPath, curName));
                }

                return returnList;
            }
        }

        /// <summary>
        ///     Gets the version of the running instance of SmartPlant Review.
        /// </summary>
        public string Version
        {
            get
            {
                // Return null if not connected
                if (!IsConnected) return null;

                // Create the parameter
                var vers = string.Empty;

                // Get the version of the SPR Application
                DrApi.Version(ref vers);

                // Return the version
                return vers;
            }
        }

        /// <summary>
        ///     Gets the next tag number to be used.
        /// </summary>
        public int NextTag
        {
            get
            {
                // Return -1 if not connected
                if (!IsConnected) return -1;

                // Create the parameter
                int returnTag;

                // Get the next tag number
                DrApi.TagNextNumber(out returnTag, 0);

                // Return the next tag number
                return returnTag;
            }
        }

        /// <summary>
        ///     Gets the next annotation number to be used.
        /// </summary>
        public int NextAnnotation
        {
            get
            {
                // Return -1 if SPR isn't running
                if (!IsConnected) return -1;

                var tbl_Site = DbMethods.GetDbTable(MdbPath, "site_table");
                return Convert.ToInt32(tbl_Site.Rows[0]["next_text_anno_id"]);
            }
        }

        /// <summary>
        ///     Determines if the SmartPlant Review application is busy.
        /// </summary>
        public bool IsBusy
        {
            get
            {
                // Return false if not connected
                if (!IsConnected) return false;

                // Get the SPR Application process
                uint procId;
                DrApi.ProcessIdGet(out procId);

                // Get a starting point to measure
                var sw = new Stopwatch();

                // Wait until the application is idle (10 ms max)
                sw.Start();
                Process.GetProcessById((int)procId).WaitForInputIdle(10);
                sw.Stop();

                // Return true if the application waited the full test period
                return (sw.ElapsedMilliseconds >= 10);
            }
        }

        /// <summary>
        ///     Gets the process ID of the SmartPlant Review application.
        /// </summary>
        public IntPtr ProcessId
        {
            get
            {
                // Return zero if not conencted
                if (!IsConnected) return IntPtr.Zero;

                uint procId;
                DrApi.ProcessIdGet(out procId);

                return (IntPtr)procId;
            }
        }

        /// <summary>
        ///     The primary SmartPlant Review application window.
        /// </summary>
        public SprWindow ApplicationWindow
        {
            get { return IsConnected ? Window_Get(SprConstants.SprApplicationWindow) : null; }
            set { if (IsConnected) Window_Set(value); }
        }

        /// <summary>
        ///     The Main view window inside the SmartPlant Review application.
        /// </summary>
        public SprWindow MainWindow
        {
            get { return IsConnected ? Window_Get(SprConstants.SprMainWindow) : null; }
            set { if (IsConnected) Window_Set(value); }
        }

        /// <summary>
        ///     The Plan view window inside the SmartPlant Review application.
        /// </summary>
        public SprWindow PlanWindow
        {
            get { return IsConnected ? Window_Get(SprConstants.SprPlanWindow) : null; }
            set { if (IsConnected) Window_Set(value); }
        }

        /// <summary>
        ///     The Elevation window inside the SmartPlant Review application.
        /// </summary>
        public SprWindow ElevationWindow
        {
            get { return IsConnected ? Window_Get(SprConstants.SprElevationWindow) : null; }
            set { if (IsConnected) Window_Set(value); }
        }

        /// <summary>
        ///     The Text window inside the SmartPlant Review application.
        /// </summary>
        public SprWindow TextWindow
        {
            get { return IsConnected ? Window_Get(SprConstants.SprTextWindow) : null; }
            set { if (IsConnected) Window_Set(value); }
        }

        /// <summary>
        ///     The default properties used when an SprSnapshot is omitted from snapshot methods.
        /// </summary>
        public SprSnapShot DefaultSnapshot { get; set; }

        #endregion

        /// <summary>
        ///     Creates a new SprApplication class.  Will automatically connect to a
        ///     single SmartPlant Review process if available.
        /// </summary>
        public SprApplication()
        {
            ActiveApplication = this;
            if (SprProcesses.Length == 1)
                Connect();

            DefaultSnapshot = new SprSnapShot
            {
                AntiAlias = 3,
                OutputFormat = SprSnapshotFormat.Jpg,
                AspectOn = true,
                Scale = 1
            };

            SprSnapShot.TempDirectory = Environment.GetEnvironmentVariable("TEMP");
            SprSnapShot.DefaultDirectory = Environment.SpecialFolder.MyPictures.ToString();
        }

        /// <summary>
        ///     SprApplication deconstructor/finalizer.
        /// </summary>
        ~SprApplication()
        {
            Dispose(false);
        }

        #region General

        /// <summary>
        ///     Brings the SmartPlant Review application to the foreground.
        /// </summary>
        public void Activate()
        {
            NativeWin32.SetForegroundWindow(ApplicationWindow.WindowHandle);
        }

        /// <summary>
        ///     Connects to a running instance of SmartPlant Review.
        /// </summary>
        /// <returns>Boolean indicating success or failure of the operation.</returns>
        public bool Connect()
        {
            DrApi = null;
            try
            {
                DrApi = Activator.CreateInstance(SprImportedTypes.DrApi);
                return IsConnected;
            }
            catch{ }
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
            if (!_disposed)
            {
                if (disposing)
                {
                    if (DrApi != null)
                        Marshal.ReleaseComObject(DrApi);
                }

                _disposed = true;
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
            int sprResult = DrApi.SessionAttach(fileName);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Exports the current session to a VUE file format.
        ///     Compatible only with SPR versions 9 and above.
        /// </summary>
        /// <param name="vueName">The full path of the vue file to be exported.</param>
        public void Export(string vueName)
        {
            // Check version compatibility
            int vers = int.Parse(Version.Substring(0,2));
            if (vers < 9)
                throw SprExceptions.SprUnsupportedMethod;

            int sprResult = DrApi.ExportVue(vueName, 0);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Exits the SmartPlant Review application.
        /// </summary>
        public void Exit()
        {
            int sprResult = DrApi.ExitViewer();
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Refreshes the session data as if selected from the application menu.
        /// </summary>
        public void RefreshData()
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            Activate();
            SendKeys.SendWait("%TR");
        }

        /// <summary>
        ///     Sets the value of a specified global option in SmartPlant Review.
        /// </summary>
        /// <param name="option">The integer of the global option to set.</param>
        /// <param name="value">The integer value to set the global option to.</param>
        public void GlobalOptionsSet(int option, int value)
        {
            int sprResult = DrApi.GlobalOptionsSet(option, value);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Gets the value of a specified global option in SmartPlant Review.
        /// </summary>
        /// <param name="option">The integer value of the option to retrieve.</param>
        /// <returns>Integer value representing the current state of the option.</returns>
        public double GlobalOptionsGet(int option)
        {
            double returnVal;
            int sprResult;

            sprResult = DrApi.GlobalOptionsGet(option, out returnVal);
            SprUtilities.ErrorHandler(sprResult);
            return returnVal;
        }

        /// <summary>
        ///     Highlights a specified object in the main SmartPlant Review application Window
        /// </summary>
        /// <param name="objectId">Object Id of the of the entity to be highlighted.</param>
        public void HighlightObject(int objectId)
        {
            int sprResult = DrApi.HighlightObject(objectId, 1, 0);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Clears all highlighting from the main SmartPlant Review application window.
        /// </summary>
        public void HighlightClear()
        {
            int sprResult = DrApi.HighlightExit(1);
            SprUtilities.ErrorHandler(sprResult);

            sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
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
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var returnPoint = new SprPoint3D();
            int abortFlag;

            Activate();

            // Prompt the user for a 3D point inside SPR
            int sprResult = DrApi.PointLocateDbl(prompt, out abortFlag, ref returnPoint.DrPointDbl);
            SprUtilities.ErrorHandler(sprResult);

            // Return null if the locate operation was aborted
            if (abortFlag != 0) return null;
            return returnPoint;
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
            // Throw an error if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var returnPoint = new SprPoint3D();
            int abort;
            int objId;
            const int flag = 0;

            Activate();

            // Prompt the user for a 3D point inside SPR
            int sprResult = DrApi.PointLocateExtendedDbl(prompt, out abort, ref returnPoint.DrPointDbl,
                                                         ref targetPoint.DrPointDbl, out objId, flag);
            SprUtilities.ErrorHandler(sprResult);

            if (abort != 0) return null;
            return returnPoint;
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
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            const int filterFlag = 0;
            var returnId = -1;

            Activate();

            // Prompt the user for a 3D point inside SPR
            int sprResult = DrApi.ObjectLocateDbl(prompt, filterFlag, out returnId, ref refPoint.DrPointDbl);
            SprUtilities.ErrorHandler(sprResult);

            return returnId;
        }

        /// <summary>
        ///     Uses an Object Id to retrieve detailed object information.
        /// </summary>
        /// <param name="objectId">The ObjectId that is looked up inside SmartPlant Review.</param>
        /// <returns>SprObjectData object containing the retrieved information.</returns>
        public SprObjectData GetObjectData(int objectId)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var returnData = new SprObjectData();
            returnData.ObjectId = objectId;
            int sprResult = DrApi.ObjectDataGetDbl(objectId, 2, ref returnData.DrObjectDataDbl);
            SprUtilities.ErrorHandler(sprResult);

            // Build the data label collection
            string lblName = string.Empty, lblValue = string.Empty;
            for (var i = 0; i < returnData.DrObjectDataDbl.LabelDataCount; i++)
            {
                sprResult = DrApi.ObjectDataLabelGet(ref lblName, ref lblValue, i);
                SprUtilities.ErrorHandler(sprResult);

                // Check if the label already exists
                if (!returnData.LabelData.ContainsKey(lblName))
                    returnData.LabelData.Add(lblName, lblValue);
            }

            return returnData;
        }

        /// <summary>
        ///     Prompts a user to select an object inside SmartPlant Review.
        ///     Retrieves object information from the selected object.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <param name="singleObjects">Indicates if SmartPlant Review locates grouped objects individually.</param>
        /// <returns>The SprObjectData object containing the retrieved information.</returns>
        public SprObjectData GetObjectData(string prompt, bool singleObjects = false)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var objId = GetObjectId(prompt);
            return GetObjectData(objId);
        }

        #endregion

        #region Text Window

        /// <summary>
        ///     Clears the contents of the SmartPlant Review text window.
        /// </summary>
        public void TextWindow_Clear()
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // Send a blank string to the application text window
            int sprResult = DrApi.TextWindow(SprConstants.SprClearTextWindow, "Text View", string.Empty, 0);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Updates the contents of the SmartPlant Review text window.
        /// </summary>
        /// <param name="mainText">String to be displayed in the text window.</param>
        public void TextWindow_Update(string mainText)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var existTitle = TextWindow_GetTitle();
            TextWindow_Update(mainText, existTitle);
        }

        /// <summary>
        ///     Updates the title and contents of the SmartPlant Review text window.
        /// </summary>
        /// <param name="mainText">String to be displayed in the text window.</param>
        /// <param name="titleText">String to be displayed in the title.</param>
        public void TextWindow_Update(string mainText, string titleText)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // Set the text window and title contents
            int sprResult = DrApi.TextWindow(SprConstants.SprClearTextWindow, titleText, mainText, 0);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Gets the existing title string of the SmartPlant Review text window.
        /// </summary>
        /// <returns>The string containing the title string.</returns>
        public string TextWindow_GetTitle()
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            int sprResult = DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);
            SprUtilities.ErrorHandler(sprResult);
            
            return orgTitle ?? (string.Empty);
        }

        /// <summary>
        ///     Gets the existing contents of the SmartPlant Review text window.
        /// </summary>
        /// <returns>The string containing the text window contents.</returns>
        public string TextWindow_GetText()
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            int sprResult = DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);
            SprUtilities.ErrorHandler(sprResult);
            
            return orgText ?? (string.Empty);
        }

        #endregion

        #region Tags

        /// <summary>
        ///     Adds a tag as a new row in the Mdb tag_data table.
        /// </summary>
        /// <param name="tag">The Tag to be written to the database.</param>
        public void Tags_Add(SprTag tag)
        {
            // Get tbl_Tags
            // Add new row
            // Update MDB
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Creates a new data field in the Mdb tag_data table.
        ///     Returns true if the field already exists.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <returns>Indicates the success or failure of the table modification.</returns>
        public bool Tags_AddDataField(string fieldName)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            return DbMethods.AddDbField(MdbPath, "tag_data", fieldName);
        }

        /// <summary>
        ///     Deletes a tag from the active SmartPlant Review session.
        /// </summary>
        /// <param name="tagNo">Integer representing the tag number to delete.</param>
        /// <param name="setAsNextTag">Determines if the tag number deleted is set as the next available tag number.</param>
        public void Tags_Delete(int tagNo, bool setAsNextTag = false)
        {
            int sprResult = DrApi.TagDelete(tagNo, 0);
            SprUtilities.ErrorHandler(sprResult);

            if (setAsNextTag)
                Tags_SetNextTag(tagNo);

            sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Toggles tag display in the active SmartPlant Review Session.
        /// </summary>
        /// <param name="displayState">Determines the tag visibility state.</param>
        public void Tags_Display(SprTagVisibility displayState)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            TextWindow_Clear();
            Activate();

            // Get the menu alias character from the enumerator
            var alias = Char.ConvertFromUtf32((int)displayState);
            SendKeys.SendWait(string.Format("%GS{0}", alias));
        }

        /// <summary>
        ///     Sets the next_tag_id in the Mdb site_table 1 above the largest existing tag.
        ///     If no tags exist, the next_tag_id is set to 1.
        /// </summary>
        public void Tags_SetNextTag()
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var tbl_Site = DbMethods.GetDbTable(MdbPath, "site_table");
            var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data");
            
            if (tbl_Tags.Rows.Count > 0)

                // Set the next tag to the highest tag value + 1
                tbl_Site.Rows[0]["next_tag_id"] =
                    Convert.ToInt32(tbl_Tags.Rows[tbl_Tags.Rows.Count - 1]["tag_unique_id"]) + 1;
            else
                tbl_Site.Rows[0]["next_tag_id"] = 1;

            DbMethods.UpdateDbTable(MdbPath, tbl_Site);
        }

        /// <summary>
        ///     Sets the next_tag_id in the Mdb site_table to the specified value.
        /// </summary>
        /// <param name="tagNo">Integer of the new next_tag_id value.</param>
        public void Tags_SetNextTag(int tagNo)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var tbl_Site = DbMethods.GetDbTable(MdbPath, "site_table");
            tbl_Site.Rows[0]["next_tag_id"] = tagNo;
                
            DbMethods.UpdateDbTable(MdbPath, tbl_Site);
        }

        /// <summary>
        ///     Locates the specified tag in the SmartPlant Review application main window.
        /// </summary>
        /// <param name="tagNo">Integer of the tag number.</param>
        /// <param name="displayTag">Indicates if the tag will be displayed.</param>
        public void Tags_Goto(int tagNo, bool displayTag = true)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var curTag = Tags_Get(tagNo);
                
            // Update the text window with the tag contents
            TextWindow_Update(curTag.Text, string.Format("Tag {0}", tagNo));

            int sprResult = DrApi.GotoTag(tagNo, 0, Convert.ToInt32(displayTag));
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Retrieves the desired tag from the Mdb tag_data table.
        /// </summary>
        /// <param name="tagNo">Integer of the tag to retrieve.</param>
        /// <returns>SprTag containing the returned tag information.</returns>
        public SprTag Tags_Get(int tagNo)
        {
            var returnTag = new SprTag();
            var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data");
            var tagRow = tbl_Tags.Select(string.Format("tag_unique_id = '{0}'", tagNo))[0];
            if (tagRow == null) throw SprExceptions.SprTagNotFound;

            return SprUtilities.BuildTagFromData(tagRow);
        }

        /// <summary>
        ///     Returns a list of all SprTags currently in the MDB database.
        /// </summary>
        /// <returns>The SprTag collection.</returns>
        public List<SprTag> Tags_GetAll()
        {
            var returnList = new List<SprTag>();
            var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data").Copy();
            
            foreach (DataRow tagRow in tbl_Tags.Rows)
                returnList.Add(SprUtilities.BuildTagFromData(tagRow));

            return returnList;
        }

        /// <summary>
        ///     Prompts a user to place a tag in the SmartPlant Review main view.
        /// </summary>
        /// <param name="tagText">String containing the tag text.</param>
        public void Tags_Place(string tagText)
        {
            var tag = new SprTag { Text = tagText };
            Tags_Place(ref tag);
        }

        /// <summary>
        ///     Prompts a user to place a tag in the SmartPlant Review main view.
        /// </summary>
        /// <param name="tag">SprTag containing the tag information.</param>
        public void Tags_Place(ref SprTag tag)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var tagOrigin = new SprPoint3D();
            var objId = GetObjectId("SELECT TAG START POINT", ref tagOrigin);
            if (objId == 0)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            var tagLeader = GetPoint("SELECT TAG LEADER LOCATION", tagOrigin);
            if (tagLeader == null)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            if (objId == 0 || tagLeader == null) throw SprExceptions.SprNullPoint;

            var currentObject = GetObjectData(objId);
            dynamic tagLabelKey = currentObject.DrObjectDataDbl.LabelKey;

            // Turn label tracking on on the flag bitmask
            tag.Flags |= SprConstants.SprTagLabel;

            SprUtilities.SetTagRegistry(tag);
            int sprResult = DrApi.TagSetDbl(tag.Id, 0, tag.Flags, ref tagLeader.DrPointDbl,
                                            ref tagOrigin.DrPointDbl, tagLabelKey, tag.Text);
            SprUtilities.ErrorHandler(sprResult);

            SprUtilities.ClearTagRegistry();
            TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.Id));
        }

        /// <summary>
        ///     Prompts a user to select new leader points for an existing tag.
        /// </summary>
        /// <param name="tagNo">Integer of the tag to edit.</param>
        public void Tags_EditLeader(int tagNo)
        {
            var tag = Tags_Get(tagNo);
            Tags_EditLeader(ref tag);
        }

        /// <summary>
        ///     Prompts a user to select new leader points for an existing tag.
        /// </summary>
        /// <param name="tag">SprTag containing the tag information.</param>
        public void Tags_EditLeader(ref SprTag tag)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            if (!tag.IsPlaced) throw SprExceptions.SprTagNotPlaced;

            var tagText = tag.Text;
            var tagOrigin = new SprPoint3D();

            var objId = GetObjectId("SELECT NEW TAG START POINT", ref tagOrigin);
            if (objId == 0)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            var tagLeader = GetPoint("SELECT NEW LEADER LOCATION", tagOrigin);
            if (tagLeader == null)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            if (objId == 0 || tagLeader == null) throw SprExceptions.SprNullPoint;
            var currentObject = GetObjectData(objId);
            dynamic tagLabelKey = currentObject.DrObjectDataDbl.LabelKey;

            // Bitmask set labels/edit true
            tag.Flags |= SprConstants.SprTagLabel;
            tag.Flags |= SprConstants.SprTagEdit;

            int sprResult = DrApi.TagSetDbl(tag.Id, 0, tag.Flags, tagLeader.DrPointDbl,
                                                tagOrigin.DrPointDbl, tagLabelKey, tagText);
            SprUtilities.ErrorHandler(sprResult);

            // Flip the tag 180 degrees.  Intergraph is AWESOME!
            tag = Tags_Get(tag.Id);
            var newOrigin = tag.LeaderPoint;
            var newLeader = tag.OriginPoint;
            tag.LeaderPoint = newLeader;
            tag.OriginPoint = newOrigin;
            Tags_Update(tag);

            TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.Id));
            sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Updates tag information directly in the Mdb tag_data table.
        /// </summary>
        /// <param name="tag">SprTag containing the tag information.</param>
        /// <returns>Indicates the success or failure of the tag_table modification.</returns>
        public bool Tags_Update(SprTag tag)
        {
            var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data");
            var rowFilter = string.Format("tag_unique_id = {0}", tag.Id);
            var tagRow = tbl_Tags.Select(rowFilter)[0];

            foreach (var kvp in tag.Data)
                tagRow[kvp.Key] = kvp.Value;
            
            return DbMethods.UpdateDbTable(MdbPath, tagRow);
        }

        /// <summary>
        ///     Updates tag information directly in the Mdb tag_data table, queueable from the threadpool.
        /// </summary>
        /// <param name="stateInfo">SprTag passed as an object per WaitCallback requirements.</param>
        public void Tags_Update(object stateInfo)
        {
            var tag = stateInfo as SprTag;
            if (tag == null) return;
            Tags_Update(tag);
        }

        /// <summary>
        ///     Saves a tag image in the default snapshot format as a block of binary data inside the Mdb.
        /// </summary>
        /// <param name="tagNo">The tag Id the image will be linked to.</param>
        /// <param name="ZoomToTag">Determines if the main view zooms to the tag in the main SmartPlant screen.</param>
        /// <returns></returns>
        public bool Tags_SaveImageToMDB(int tagNo, bool ZoomToTag)
        {
            return Tags_SaveImageToMDB(tagNo, DefaultSnapshot, ZoomToTag);
        }

        /// <summary>
        ///     Saves a tag image as a block of binary data inside the Mdb.
        /// </summary>
        /// <param name="tagNo">The tag Id the image will be linked to.</param>
        /// <param name="snap">The snapshot format the image will be created with.</param>
        /// <param name="ZoomToTag">Determines if the main view zooms to the tag in the main SmartPlant screen.</param>
        /// <returns></returns>
        public bool Tags_SaveImageToMDB(int tagNo, SprSnapShot snap, bool ZoomToTag)
        {
            if (ZoomToTag) Tags_Goto(tagNo);
            var imgPath = TakeSnapshot(snap, "dbImage_temp", SprSnapShot.TempDirectory);

            if (!DbMethods.AddDbField(MdbPath, "tag_data", "tag_image", "OLEOBJECT"))
                return false;

            using (var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
            {
                var imgBytes = new byte[fs.Length];
                fs.Read(imgBytes, 0, imgBytes.Length);

                var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data");
                var rowFilter = string.Format("tag_unique_id = {0}", tagNo);
                var tagRow = tbl_Tags.Select(rowFilter)[0];
                tagRow["tag_image"] = imgBytes;

                // Return the result of the table update
                if (!DbMethods.UpdateDbTable(MdbPath, tagRow))
                    return false;
            }

            File.Delete(imgPath);
            return true;
        }
        
        /// <summary>
        ///     Saves images in the default snapshot format for all existing tags in the Mdb.
        /// </summary>
        public void Tags_SaveAllImagesToMDB()
        {
            Tags_SaveAllImagesToMDB(DefaultSnapshot);
        }

        /// <summary>
        ///     Saves images for all existing tags in the Mdb.
        /// </summary>
        /// <param name="snap">The snapshot format the images will be created with.</param>
        public void Tags_SaveAllImagesToMDB(SprSnapShot snap)
        {
            var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data");
            foreach (DataRow tagRow in tbl_Tags.Rows)
                Tags_SaveImageToMDB(Convert.ToInt32(tagRow["tag_unique_id"]), snap, true);
        }

        /// <summary>
        ///     Saves images in the default snapshot format for all existing tags to a local directory.
        /// </summary>
        /// <param name="nameFormat">The naming format the tags will be saved using. (## represents the tag number)</param>
        /// <param name="outputDir">The path to the directory where the images will be placed.</param>
        /// <returns></returns>
        public bool Tags_TakeSnapshots(string nameFormat, string outputDir)
        {
            return Tags_TakeSnapshots(nameFormat, outputDir, DefaultSnapshot);
        }

        /// <summary>
        ///     Saves images for all existing tags to a local directory.
        /// </summary>
        /// <param name="nameFormat">The naming format the tags will be saved using. (## represents the tag number)</param>
        /// <param name="outputDir">The path to the directory where the images will be placed.</param>
        /// <param name="snap">The snapshot format the images will be created with.</param>
        /// <returns></returns>
        public bool Tags_TakeSnapshots(string nameFormat, string outputDir, SprSnapShot snap)
        {
            // Uses the unique_tag_id to auto-increment
            nameFormat = nameFormat.Replace("##", "{0}");

            var tbl_Tags = DbMethods.GetDbTable(MdbPath, "tag_data");
            for (int i = 0; i < tbl_Tags.Rows.Count; i++)
            {
                var curTagNo = Convert.ToInt32(tbl_Tags.Rows[i]["tag_unique_id"]);
                Tags_Goto(curTagNo);

                TakeSnapshot(snap, string.Format(nameFormat, curTagNo), outputDir);
            }

            return true;
        }

        #endregion

        #region Annotation

        /// <summary>
        ///     Toggles annotation display in the SmartPlant Review application main window.
        /// </summary>
        /// <param name="visible">Determines the annotation visibility state.</param>
        public void Annotations_Display(bool visible)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var visValue = Convert.ToInt32(visible);
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            if (objViewdataDbl == null) throw SprExceptions.SprObjectCreateFail;

            int sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);
            SprUtilities.ErrorHandler(sprResult);
            objViewdataDbl.AllAnnotationsDisplay = visValue;

            // Update the global annotation visibility properties
            sprResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDisplay, visValue);
            SprUtilities.ErrorHandler(sprResult);

            sprResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoTextDisplay, visValue);
            SprUtilities.ErrorHandler(sprResult);

            sprResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDataDisplay, visValue);
            SprUtilities.ErrorHandler(sprResult);
                        
            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Creates a new data field in the Mdb text_annotations table.
        ///     Returns true if the field already exists.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <returns>Indicates the success or failure of the table modification.</returns>
        public bool Annotations_AddDataField(string fieldName)
        {
            return DbMethods.AddDbField(MdbPath, "text_annotations", fieldName);
        }

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
                TextWindow_Update("Annotation placement canceled.");
                return;
            }        

            // Get the annotation center point using the leaderpoint for depth calculation
            var centerPoint = GetPoint("SELECT THE CENTER POINT FOR THE ANNOTATION LABEL", leaderPoint);

            // Exit if the centerpoint selection failed/canceled
            if (centerPoint == null)
            {
                TextWindow_Update("Annotation placement canceled.");
                return;
            }

            // Create a reference to the DrAnnotation
            var drAnno = anno.DrAnnotationDbl;

            // Set the annotation points
            drAnno.LeaderPoint = leaderPoint.DrPointDbl;
            drAnno.CenterPoint = centerPoint.DrPointDbl;

            // Place the annotation on screen
            int annoId;
            int sprResult = DrApi.AnnotationCreateDbl(anno.Type, ref drAnno, out annoId);
            SprUtilities.ErrorHandler(sprResult);

            // Link the located object to the annotation
            sprResult = DrApi.AnnotationDataSet(annoId, anno.Type, ref drAnno, ref objId);
            SprUtilities.ErrorHandler(sprResult);

            // Retrieve the placed annotation data
            anno = Annotations_Get(anno.Id);

            // Add an ObjectId field
            Annotations_AddDataField("object_id");

            // Save the ObjectId to the annotation data
            anno.Data["object_id"] = objId;

            // Update the annotation
            Annotations_Update(anno);

            // Update the text window
            TextWindow_Update(anno.Text, string.Format("Annotation {0}", anno.Id));

            // Update the main view
            sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
        }

        public void Annotations_EditLeader(int annoNo)
        {
            // Update the placement points and reload?
        }
        public void Annotations_EditLeader(ref SprAnnotation annotation)
        {
            // Update the placement points and reload?
        }

        /// <summary>
        ///     Adds an annotation directly into the Mdb annotation table.
        /// </summary>
        /// <param name="anno">SprAnnotation containing the annotation information.</param>
        public void Annotations_Add(SprAnnotation anno)
        {
            // Get the tbl_Annotations
            // create a new row with the anno values
            // update the MDB
        }

        /// <summary>
        ///     Adds a list of annotations into the Mdb annotation table.
        /// </summary>
        /// <param name="annotations">List of SprAnnotation to be added to the annotation table.</param>
        public void Annotations_Add(List<SprAnnotation> annotations)
        {
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
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            int selectedId;
            int assocId;
            var annotation = new SprAnnotation();

            Activate();

            var promptString = string.Format("SELECT THE DESIRED {0} ANNOTATION", type.ToUpper());
            int sprResult = DrApi.AnnotationLocate(type, promptString, 0, out selectedId);
            SprUtilities.ErrorHandler(sprResult);

            annotation.Id = selectedId;
            
            var drAnno = annotation.DrAnnotationDbl;
            sprResult = DrApi.AnnotationDataGet(selectedId, type, ref drAnno, out assocId);
            SprUtilities.ErrorHandler(sprResult);
            
            if (assocId == 0) return null;
            annotation.AssociatedObject = GetObjectData(assocId);

            return annotation;
        }

        /// <summary>
        ///     Prompts a user to delete an annotation in the SmartPlant Review main view.
        /// </summary>
        /// <param name="type">The string type of the annotation to delete.</param>
        public void Annotations_Delete(string type)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            int annoId;

            Activate();

            var promptString = string.Format("SELECT THE {0} ANNOTATION TO DELETE", type.ToUpper());
            int sprResult = DrApi.AnnotationLocate(type, promptString, 0, out annoId);
            SprUtilities.ErrorHandler(sprResult);

            if (annoId == 0) return;
            sprResult = DrApi.AnnotationDelete(type, annoId, 0);
            SprUtilities.ErrorHandler(sprResult);

            sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Deletes all the annotations of a specified type from the active SmartPlant Review session.
        /// </summary>
        /// <param name="type">The string type of the annotations to delete.</param>
        public void Annotations_DeleteType(string type)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            int sprResult = DrApi.AnnotationDeleteAll(type, 0);
            SprUtilities.ErrorHandler(sprResult);

            sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Deletes all the annotations from the active SmartPlant Review session.
        /// </summary>
        public void Annotations_DeleteAll()
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            var tbl_TextAnnotations = DbMethods.GetDbTable(MdbPath, "text_annotation_types");
            for (var i = tbl_TextAnnotations.Rows.Count - 1; i >= 0; i--)
                Annotations_DeleteType(tbl_TextAnnotations.Rows[i]["name"].ToString());

            int sprResult = DrApi.ViewUpdate(1);
            SprUtilities.ErrorHandler(sprResult);
        }

        /// <summary>
        ///     Retrieves the desired annotation from the Mdb text_annotations table.
        /// </summary>
        /// <param name="annoId">Unique Id of the annotation to retrieve.</param>
        /// <returns>SprAnnotation containing the retirned annotation information.</returns>
        public SprAnnotation Annotations_Get(int annoId)
        {
            var annotation = new SprAnnotation();
            var tbl_Annotations = DbMethods.GetDbTable(MdbPath, "text_annotations");
            var annotationRow = tbl_Annotations.Select(string.Format("id = '{0}'", annoId))[0];

            foreach (DataColumn col in tbl_Annotations.Columns)
            {
                annotation.Data[col.ColumnName] = annotationRow[col];
            }

            annotation.IsPlaced = true;
            return annotation;
        }

        /// <summary>
        ///     Updates annotation data directly in the Mdb text_annotations table.
        /// </summary>
        /// <param name="anno">SprAnnotation containing the annotation information.</param>
        /// <returns>Indicates the success or failure of the text_annotations modification.</returns>
        public bool Annotations_Update(SprAnnotation anno)
        {
            var tbl_Annotations = DbMethods.GetDbTable(MdbPath, "text_annotations");
            var rowFilter = string.Format("id = {0}", anno.Id);
            var annotationRow = tbl_Annotations.Select(rowFilter)[0];

            foreach (var kvp in anno.Data)
                annotationRow[kvp.Key] = kvp.Value;

            return DbMethods.UpdateDbTable(MdbPath, annotationRow);
        }

        /// <summary>
        ///     Updates annotation information directly in the Mdb text_annotations table, queueable from the threadpool.
        /// </summary>
        /// <param name="stateInfo">SprAnnotation passed as an object per WaitCallback requirements.</param>
        public void Annotations_Update(object stateInfo)
        {
            var annotation = stateInfo as SprAnnotation;
            if (annotation == null) return;

            Annotations_Update(annotation);
        }

        #endregion

        #region Snapshot

        public string TakeSnapshot(string imageName, string outputDir)
        {
            return TakeSnapshot(DefaultSnapshot, imageName, outputDir);
        }
        /// <summary>
        ///     Captures the active SmartPlant Reviews session main view and 
        ///     writes it an image of the given filename and format.
        /// </summary>
        /// <param name="snapShot">SprSnapshot containing the snapshot settings.</param>
        /// <param name="imageName">Name of the final output image.</param>
        /// <param name="outputDir">Directory to save the snapshot to.</param>
        /// <returns></returns>
        public string TakeSnapshot(SprSnapShot snapShot, string imageName, string outputDir)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // .BMP is forced before conversions
            var imgPath = Path.Combine(outputDir, string.Format("{0}.bmp", imageName));

            // Turn on view backfaces/endcaps
            var orgBackfaces = GlobalOptionsGet(SprConstants.SprGlobalBackfacesDisplay);
            var orgEndcaps = GlobalOptionsGet(SprConstants.SprGlobalEndcapsDisplay);
            if (orgBackfaces == 0)
            GlobalOptionsSet(SprConstants.SprGlobalBackfacesDisplay, 1);
            if (orgEndcaps == 0)
            GlobalOptionsSet(SprConstants.SprGlobalEndcapsDisplay, 1);

            int sprResult = DrApi.SnapShot(imgPath, snapShot.Flags, snapShot.DrSnapShot, 0);
            while (IsBusy) Thread.Sleep(100);
            SprUtilities.ErrorHandler(sprResult);

            if (orgBackfaces == 0) GlobalOptionsSet(SprConstants.SprGlobalBackfacesDisplay, 1);
            if (orgEndcaps == 0) GlobalOptionsSet(SprConstants.SprGlobalEndcapsDisplay, 1);

            if (!File.Exists(imgPath)) return string.Empty;

            if (snapShot.OutputFormat != SprSnapshotFormat.Bmp)
                imgPath = SprSnapShot.FormatSnapshot(imgPath, snapShot.OutputFormat);

            return imgPath;
        }
      
        public bool ExportPDF(int quality, string path)
        {
            int sprResult = DrApi.ExportPDF(path, quality, 1, 1, 1);
            SprUtilities.ErrorHandler(sprResult);

            return true;
        }

        #endregion

        #region Views

        public void SetCenterPoint(double east, double north, double elevation)
        {
            SetCenterPoint(new SprPoint3D(east, north, elevation));
        }

        public void SetCenterPoint(SprPoint3D centerPoint)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            int sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);
            SprUtilities.ErrorHandler(sprResult);

            objViewdataDbl.CenterUorPoint = centerPoint.DrPointDbl;

            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
            SprUtilities.ErrorHandler(sprResult);
        }

        public void SetEyePoint(double east, double north, double elevation)
        {
            SetEyePoint(new SprPoint3D(east, north, elevation));
        }

        public void SetEyePoint(SprPoint3D eyePoint)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);
            int sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);
            SprUtilities.ErrorHandler(sprResult);

            objViewdataDbl.EyeUorPoint = eyePoint.DrPointDbl;

            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
            SprUtilities.ErrorHandler(sprResult);
        }

        //public void GotoPoint

        #endregion

        #region Windows

        /// <summary>
        ///     Gets a SmartPlant Review application window.
        /// </summary>
        /// <param name="windowNo">The integer of the window to return.</param>
        /// <returns>The SprWindow containing the window properties.</returns>
        public SprWindow Window_Get(int windowNo)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            int drHwnd;
            var curWin = new SprWindow();
            dynamic objWin = Activator.CreateInstance(SprImportedTypes.DrWindow);

            int sprResult = DrApi.WindowGet(windowNo, out objWin);
            SprUtilities.ErrorHandler(sprResult);

            sprResult = DrApi.WindowHandleGet(windowNo, out drHwnd);
            SprUtilities.ErrorHandler(sprResult);

            if (objWin == null) return null;

            curWin.Height = objWin.Height;
            curWin.Width = objWin.Width;
            curWin.Left = objWin.Left < 0 ? 0 : objWin.Left; // 0 if
            curWin.Top = objWin.Top < 0 ? 0 : objWin.Top;    // negative
            curWin.WindowHandle = drHwnd;
            curWin.Index = windowNo;
            return curWin;
        }

        /// <summary>
        ///     Modifies a SmartPlant Review application window.
        /// </summary>
        /// <param name="window">The SprWindow containing the new properties.</param>
        public void Window_Set(SprWindow window)
        {
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            dynamic objWin = Activator.CreateInstance(SprImportedTypes.DrWindow);
            int sprResult = DrApi.WindowGet(window.Index, out objWin);
            SprUtilities.ErrorHandler(sprResult);

            if (objWin == null) return;

            objWin.Height = window.Height;
            objWin.Width = window.Width;
            if (window.Left > 0) objWin.Left = window.Left;
            if (window.Top > 0) objWin.Top = window.Top;

            sprResult = DrApi.WindowSet(window.Index, objWin);
            SprUtilities.ErrorHandler(sprResult);
        }

        #endregion
    }
}