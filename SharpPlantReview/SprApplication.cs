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
using System.Windows.Forms;
using System.Threading;

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
                LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);

                // Get the directory path
                LastResult = DrApi.FilePathFromNumber(1, ref dirPath);

                // Get the MDB FileName
                LastResult = DrApi.FileNameFromNumber(1, ref mdbName);

                // Reset the global variable
                LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

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
                string sessionName = null;

                // Set the global variable
                LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 1);

                // Get the session file name
                LastResult = DrApi.FileNameFromNumber(0, ref sessionName);

                // Reset the global variable
                LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

                // Return the path
                return sessionName;
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
                LastResult = DrApi.FileCountGet(out fileCount);

                GlobalOptionsSet(SprConstants.SprGlobalFileInfoMode, 0);

                var returnList = new List<string>();

                for (int i = 0; i < fileCount; i++)
                {
                    string curName;
                    string curPath;
                    LastResult = DrApi.FileNameFromNumber(i, out curName);
                    LastResult = DrApi.FilePathFromNumber(i, out curPath);

                    returnList.Add(Path.Combine(curPath, curName));
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
                LastResult = DrApi.Version(ref vers);

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
                LastResult = DrApi.TagNextNumber(out returnTag, 0);

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
                // Return -1 if not connected
                if (!IsConnected) return -1;

                // Create the parameter
                var returnAnno = -1;

                // Retrieve the MDB site table
                var siteTable = DbMethods.GetDbTable(MdbPath, "site_table");

                if (siteTable != null)

                    // Set the next annotation number
                    returnAnno = (int)siteTable.Rows[0]["next_text_anno_id"];

                // Return the next annotation number
                return returnAnno;
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

        /// <summary>
        ///     The returned result from the most recent DrApi function call.
        /// </summary>
        public int LastResult
        {
            get { return _lastResult; }
            internal set
            {
                _lastResult = value;

                // Handle the errors
                SprUtilities.ErrorCheck(value);
            }
        }
        private int _lastResult;

        /// <summary>
        ///     The last error message returned from the most recent DrApi function call.
        /// </summary>
        public string LastError { get; internal set; }

        #endregion

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
            // Clear the current Application
            DrApi = null;

            // Get an instance of SmartPlant Review
            DrApi = Activator.CreateInstance(SprImportedTypes.DrApi);

            // Return the connection state
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
            // Throw an error if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Try opening the file using the Api method call
            LastResult = DrApi.SessionAttach(fileName);
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
                throw SprExceptions.SprVersionIncompatibility;

            // Export the current session to VUE
            LastResult = DrApi.ExportVue(vueName, 0);
        }

        /// <summary>
        ///     Exits the SmartPlant Review application.
        /// </summary>
        public void Exit()
        {
            // Throw an error if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Exit the SPR application
            LastResult = DrApi.ExitViewer();
        }

        /// <summary>
        ///     Refreshes the session data as if selected from the application menu.
        /// </summary>
        public void RefreshData()
        {
            // Thrown an error if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Set SPR to the front
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
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Set the global option to the provided value
            LastResult = DrApi.GlobalOptionsSet(option, value);
        }

        /// <summary>
        ///     Gets the value of a specified global option in SmartPlant Review.
        /// </summary>
        /// <param name="option">The integer value of the option to retrieve.</param>
        /// <returns>Integer value representing the current state of the option.</returns>
        public double GlobalOptionsGet(int option)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the return value
            double returnVal;

            // Get the global option value
            LastResult = DrApi.GlobalOptionsGet(option, out returnVal);

            // Return the retrieved value
            return returnVal;
        }

        /// <summary>
        ///     Highlights a specified object in the main SmartPlant Review application Window
        /// </summary>
        /// <param name="objectId">Object Id of the of the entity to be highlighted.</param>
        public void HighlightObject(int objectId)
        {
            // Highlight the object in SPR
            LastResult = DrApi.HighlightObject(objectId, 1, 0);
        }

        /// <summary>
        ///     Clears all highlighting from the main SmartPlant Review application window.
        /// </summary>
        public void HighlightClear()
        {
            // Clear the highlighting in SPR 
            LastResult = DrApi.HighlightExit(1);

            // Refresh the main window
            LastResult = DrApi.ViewUpdate(1);
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
            // Throw an error if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the return point
            var returnPoint = new SprPoint3D();

            // Create the params
            int abortFlag;

            // Set the SPR application visible
            Activate();

            // Prompt the user for a 3D point inside SPR
            LastResult = DrApi.PointLocateDbl(prompt, out abortFlag, ref returnPoint.DrPointDbl);

            // Return null if the locate operation was aborted
            if (abortFlag != 0) return null;

            // Return the new point
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

            // Create the return point
            var returnPoint = new SprPoint3D();

            // Create the params
            int abort;
            int objId;
            const int flag = 0;

            // Set the SPR application visible
            Activate();

            // Prompt the user for a 3D point inside SPR
            LastResult = DrApi.PointLocateExtendedDbl(prompt, out abort, ref returnPoint.DrPointDbl,
                                                         ref targetPoint.DrPointDbl, out objId, flag);

            // Return null if the locate operation was aborted
            if (abort != 0) return null;

            // Return the new point
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
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the params
            const int filterFlag = 0;
            var returnId = -1;

            // Set the SPR application visible
            Activate();

            // Prompt the user for a 3D point inside SPR
            LastResult = DrApi.ObjectLocateDbl(prompt, filterFlag, out returnId, ref refPoint.DrPointDbl);

            // Return the ObjectId
            return returnId;
        }

        /// <summary>
        ///     Uses an Object Id to retrieve detailed object information.
        /// </summary>
        /// <param name="objectId">The ObjectId that is looked up inside SmartPlant Review.</param>
        /// <returns>SprObjectData object containing the retrieved information.</returns>
        public SprObjectData GetObjectData(int objectId)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the return object
            var returnData = new SprObjectData();

            // Return null if the objectId is zero
            if (objectId == 0) return null;

            // Set the return object ID
            returnData.ObjectId = objectId;

            // Get the DataDbl object
            LastResult = DrApi.ObjectDataGetDbl(objectId, 2, ref returnData.DrObjectDataDbl);

            // Iterate through the labels
            string lblName = string.Empty, lblValue = string.Empty;
            for (var i = 0; i < returnData.DrObjectDataDbl.LabelDataCount; i++)
            {
                // Get the label key/value pair
                LastResult = DrApi.ObjectDataLabelGet(ref lblName, ref lblValue, i);

                // Check if the label already exists
                if (!returnData.LabelData.ContainsKey(lblName))

                    // Add the label data to the dictionary
                    returnData.LabelData.Add(lblName, lblValue);
            }

            // Return the data object
            return returnData;
        }

        /// <summary>
        ///     Prompts a user to select an object inside SmartPlant Review.
        ///     Retrieves object information from the selected object.
        /// </summary>
        /// <param name="prompt">The prompt string to be displayed in the application text window.</param>
        /// <returns>The SprObjectData object containing the retrieved information.</returns>
        public SprObjectData GetObjectData(string prompt)
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
        public SprObjectData GetObjectData(string prompt, bool singleObjects)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Get the ObjectID on screen
            var objId = GetObjectId(prompt);

            // Retrieve the Object data using the object Id;
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
            LastResult = DrApi.TextWindow(SprConstants.SprClearTextWindow, "Text View", string.Empty, 0);
        }

        /// <summary>
        ///     Updates the contents of the SmartPlant Review text window.
        /// </summary>
        /// <param name="mainText">String to be displayed in the text window.</param>
        public void TextWindow_Update(string mainText)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Get the existing title
            var existTitle = TextWindow_GetTitle();

            // Set the text window without changing the title
            TextWindow_Update(mainText, existTitle);
        }

        /// <summary>
        ///     Updates the title and contents of the SmartPlant Review text window.
        /// </summary>
        /// <param name="mainText">String to be displayed in the text window.</param>
        /// <param name="titleText">String to be displayed in the title.</param>
        public void TextWindow_Update(string mainText, string titleText)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // Set the text window and title contents
            LastResult = DrApi.TextWindow(SprConstants.SprClearTextWindow, titleText, mainText, 0);
        }

        /// <summary>
        ///     Gets the existing title string of the SmartPlant Review text window.
        /// </summary>
        /// <returns>The string containing the title string.</returns>
        public string TextWindow_GetTitle()
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the params
            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            // Get the existing text window values
            LastResult = DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Return the title, empty string if null
            return orgTitle ?? (string.Empty);
        }

        /// <summary>
        ///     Gets the existing contents of the SmartPlant Review text window.
        /// </summary>
        /// <returns>The string containing the text window contents.</returns>
        public string TextWindow_GetText()
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Params for retrieving SPR data
            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            // Get the existing text window values
            LastResult = DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Set an empty string for null values
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
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Add the tag field to the MDB database
            return DbMethods.AddDbField(MdbPath, "tag_data", fieldName);
        }

        /// <summary>
        ///     Deletes a tag from the active SmartPlant Review session.
        /// </summary>
        /// <param name="tagNo">Integer representing the tag number to delete.</param>
        public void Tags_Delete(int tagNo)
        {
            Tags_Delete(tagNo, false);
        }

        /// <summary>
        ///     Deletes a tag from the active SmartPlant Review session.
        /// </summary>
        /// <param name="tagNo">Integer representing the tag number to delete.</param>
        /// <param name="setAsNextTag">Determines if the tag number deleted is set as the next available tag number.</param>
        public void Tags_Delete(int tagNo, bool setAsNextTag)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Delete the desired tag
            LastResult = DrApi.TagDelete(tagNo, 0);

            // Set the deleted tag as the next tag number
            if (setAsNextTag)
                Tags_SetNextTag(tagNo);

            // Update the SmartPlant Review main view
            LastResult = DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Toggles tag display in the active SmartPlant Review Session.
        /// </summary>
        /// <param name="displayState">Determines the tag visibility state.</param>
        public void Tags_Display(SprTagVisibility displayState)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Clear the text window
            TextWindow_Clear();

            // Set SPR to the front
            Activate();

            // Get the menu alias character from the enumerator
            var alias = Char.ConvertFromUtf32((int)displayState);

            // Set the tag visibility
            SendKeys.SendWait(string.Format("%GS{0}", alias));
        }

        /// <summary>
        ///     Sets the next_tag_id in the Mdb site_table 1 above the largest existing tag.
        ///     If no tags exist, the next_tag_id is set to 1.
        /// </summary>
        public void Tags_SetNextTag()
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // Get the tags
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // Exit if the tag table was not retrieved
            if (tagTable == null) return;

            // Retrieve the site table
            var siteTable = DbMethods.GetDbTable(MdbPath, "site_table");

            // Exit if the site table was not retrieved
            if (siteTable == null) return;

            // If tags exist
            if (tagTable.Rows.Count > 0)

                // Set the next tag to the highest tag value + 1
                siteTable.Rows[0]["next_tag_id"] =
                    Convert.ToInt32(tagTable.Rows[tagTable.Rows.Count - 1]["tag_unique_id"]) + 1;
            else

                // Set the next tag to 1
                siteTable.Rows[0]["next_tag_id"] = 1;

            // Update the site table
            DbMethods.UpdateDbTable(MdbPath, siteTable);
        }

        /// <summary>
        ///     Sets the next_tag_id in the Mdb site_table to the specified value.
        /// </summary>
        /// <param name="tagNo">Integer of the new next_tag_id value.</param>
        public void Tags_SetNextTag(int tagNo)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Get the current database
            var siteTable = DbMethods.GetDbTable(MdbPath, "site_table");

            // Exit if the table is null
            if (siteTable == null) return;
            
            // Get the top row
            var row = siteTable.Rows[0];

            // Set the next tag value
            row["next_tag_id"] = tagNo;
                
            // Update the site table
            DbMethods.UpdateDbTable(MdbPath, siteTable);
        }

        /// <summary>
        ///     Locates the specified tag in the SmartPlant Review application main window.
        /// </summary>
        /// <param name="tagNo">Integer of the tag number.</param>
        public void Tags_Goto(int tagNo)
        {
            Tags_Goto(tagNo, true);
        }

        /// <summary>
        ///     Locates the specified tag in the SmartPlant Review application main window.
        /// </summary>
        /// <param name="tagNo">Integer of the tag number.</param>
        /// <param name="displayTag">Indicates if the tag will be displayed.</param>
        public void Tags_Goto(int tagNo, bool displayTag)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Get the tag data
            var curTag = Tags_Get(tagNo);

            // Exit if the tag was not retrieved
            if (curTag == null) return;
                
            // Update the text window with the tag information
            TextWindow_Update(curTag.Data["tag_text"].ToString(), string.Format("Tag {0}", tagNo));

            // Locate the desired tag on the main screen with the specified visibility
            LastResult = DrApi.GotoTag(tagNo, 0, Convert.ToInt32(displayTag));
        }

        /// <summary>
        ///     Retrieves the desired tag from the Mdb tag_data table.
        /// </summary>
        /// <param name="tagNo">Integer of the tag to retrieve.</param>
        /// <returns>SprTag containing the retirned tag information.</returns>
        public SprTag Tags_Get(int tagNo)
        {
            // Throw an exception if not connected
            //if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the new tag
            var returnTag = new SprTag();

            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // Return null if the table retrieval failed
            if (tagTable == null) return null;
            
            // Return null if no tags exist
            if (tagTable.Rows.Count == 0) return null;

            // Create the row filter for the desired tag
            var rowFilter = tagTable.Select(string.Format("tag_unique_id = '{0}'", tagNo));
            
            // Throw an exception if the tag was not found
            if (rowFilter.Length == 0) throw SprExceptions.SprTagNotFound;

            // Iterate through each column
            foreach (DataColumn col in tagTable.Columns)
            {
                // Add the key/value from the first filtered row to the dictionary
                returnTag.Data[col.ColumnName] = rowFilter[0][col];
            }

            // Return the tag
            return returnTag;
        }

        /// <summary>
        ///     Returns a list of all SprTags currently in the MDB database.
        /// </summary>
        /// <returns>The SprTag collection.</returns>
        public List<SprTag> Tags_GetAll()
        {
            // Create the return list
            var returnList = new List<SprTag>();

            // Get the tag table from the MDB 
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data").Copy();
            if (tagTable == null) return null;

            // Iterate through each tag in the table
            foreach (DataRow tagRow in tagTable.Rows)
            {
                // Add a new serialized tag to the return list
                returnList.Add(SprUtilities.BuildTagFromData(tagRow));
            }

            // Return the completed list
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
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the origin point
            var tagOrigin = new SprPoint3D();

            // Get an object on screen and set the origin point to its location
            var objId = GetObjectId("SELECT TAG START POINT", ref tagOrigin);

            // Exit if the object selection failed
            if (objId == 0)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            // Get the tag leader point using the origin for depth
            var tagLeader = GetPoint("SELECT TAG LEADER LOCATION", tagOrigin);

            // Exit if the leader point is not set
            if (tagLeader == null)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            // Throw an exception if either of the point retrievals failed
            if (objId == 0 || tagLeader == null) throw SprExceptions.SprNullPoint;

            // Get the current object for the label key
            var currentObject = GetObjectData(objId);
            dynamic tagLabelKey = currentObject.DrObjectDataDbl.LabelKey;

            // Turn label tracking on on the flag bitmask
            tag.Flags |= SprConstants.SprTagLabel;

            // Set the tag registry values
            SprUtilities.SetTagRegistry(tag);

            // Place the tag
            LastResult = DrApi.TagSetDbl(tag.Id, 0, tag.Flags, ref tagLeader.DrPointDbl,
                                            ref tagOrigin.DrPointDbl, tagLabelKey, tag.Text);

            // Retrieve the placed tag data
            tag = Tags_Get(tag.Id);

            // Clear the tag registry
            SprUtilities.ClearTagRegistry();

            // Update the text window
            TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.Id));
        }

        /// <summary>
        ///     Prompts a user to select new leader points for an existing tag.
        /// </summary>
        /// <param name="tagNo">Integer of the tag to edit.</param>
        public void Tags_EditLeader(int tagNo)
        {
            // Get the existing tag
            var tag = Tags_Get(tagNo);

            // Edit the tag leader
            Tags_EditLeader(ref tag);
        }

        /// <summary>
        ///     Prompts a user to select new leader points for an existing tag.
        /// </summary>
        /// <param name="tag">SprTag containing the tag information.</param>
        public void Tags_EditLeader(ref SprTag tag)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Throw an exception if the tag is not placed
            if (!tag.IsPlaced) throw SprExceptions.SprTagNotPlaced;

            // Get the existing tag text
            var tagText = tag.Text;

            // Create the origin point
            var tagOrigin = new SprPoint3D();

            // Get an object on screen and set the origin point to its location
            var objId = GetObjectId("SELECT NEW TAG START POINT", ref tagOrigin);

            // Exit if the object selection failed
            if (objId == 0)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            // Get the tag leader point
            var tagLeader = GetPoint("SELECT NEW LEADER LOCATION", tagOrigin);

            // Exit if the leader point is not set
            if (tagLeader == null)
            {
                TextWindow_Update("Tag placement canceled.");
                return;
            }

            // Throw an exception if either of the point retrievals failed
            if (objId == 0 || tagLeader == null) throw SprExceptions.SprNullPoint;

            // Get the current object for the label key
            var currentObject = GetObjectData(objId);
            dynamic tagLabelKey = currentObject.DrObjectDataDbl.LabelKey;

            // Turn label tracking on on the flag bitmask
            tag.Flags |= SprConstants.SprTagLabel;

            // Set the edit flag on the existing tag
            tag.Flags |= SprConstants.SprTagEdit;

            // Update the tag with the new leader points
            LastResult = DrApi.TagSetDbl(tag.Id, 0, tag.Flags, tagLeader.DrPointDbl,
                                                tagOrigin.DrPointDbl, tagLabelKey, tagText);

            // Reference the placed tag
            tag = Tags_Get(tag.Id);

            // Flip the tag 180 degrees.  Intergraph is AWESOME!
            var newOrigin = tag.LeaderPoint;
            var newLeader = tag.OriginPoint;
            tag.LeaderPoint = newLeader;
            tag.OriginPoint = newOrigin;

            // Update the tag
            Tags_Update(tag);

            // Update the text window
            TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.Id));

            // Update the main view
            LastResult = DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Updates tag information directly in the Mdb tag_data table.
        /// </summary>
        /// <param name="tag">SprTag containing the tag information.</param>
        /// <returns>Indicates the success or failure of the tag_table modification.</returns>
        public bool Tags_Update(SprTag tag)
        {
            // Throw an exception if not connected
            //if (!IsConnected)throw SprExceptions.SprNotConnected;

            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // Return false if the table is null
            if (tagTable == null) return false;

            // Return false if no tags exist
            if (tagTable.Rows.Count == 0) return false;
            
            // Create the row filter for the specified tag
            var rowFilter = string.Format("tag_unique_id = {0}", tag.Id);
            var tblFilter = tagTable.Select(rowFilter);

            // Iterate through each dictionary key/value pair
            foreach (var kvp in tag.Data)
            
            // Set the values for the selected tag
            tblFilter[0][kvp.Key] = kvp.Value;
            
            // Return the result of the table update
            return DbMethods.UpdateDbTable(MdbPath, rowFilter, tagTable);
        }

        /// <summary>
        ///     Updates tag information directly in the Mdb tag_data table, queueable from the threadpool.
        /// </summary>
        /// <param name="stateInfo">SprTag passed as an object per WaitCallback requirements.</param>
        public void Tags_Update(object stateInfo)
        {
            // Cast the threading object 
            var tag = stateInfo as SprTag;

            // Return if the tag is null
            if (tag == null) return;

            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // Return if the table is null
            if (tagTable == null) return;

            // Return if no tags exist
            if (tagTable.Rows.Count == 0) return;

            // Create the row filter for the specified tag
            var rowFilter = string.Format("tag_unique_id = {0}", tag.Id);
            var tblFilter = tagTable.Select(rowFilter);

            // Iterate through each dictionary key/value pair
            foreach (var kvp in tag.Data)

                // Set the values for the selected tag
                tblFilter[0][kvp.Key] = kvp.Value;

            // Push the the updated table
            DbMethods.UpdateDbTable(MdbPath, rowFilter, tagTable);
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
            // Zoom to the tag as needed
            if (ZoomToTag)
                Tags_Goto(tagNo);

            string imgPath = TakeSnapshot(snap, "dbImage_temp", SprSnapShot.TempDirectory);

            if (!DbMethods.AddDbField(MdbPath, "tag_data", "tag_image", "OLEOBJECT"))
                return false;

            using (var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
            {
                var imgBytes = new byte[fs.Length];
                fs.Read(imgBytes, 0, imgBytes.Length);

                var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

                var rowFilter = string.Format("tag_unique_id = {0}", tagNo);
                var tblFilter = tagTable.Select(rowFilter);
                tblFilter[0]["tag_image"] = imgBytes;

                // Return the result of the table update
                if (!DbMethods.UpdateDbTable(MdbPath, rowFilter, tagTable))
                    return false;
            }

            File.Delete(imgPath);
            return true;
        }
        
        /// <summary>
        ///     Saves images in the default snapshot format for all existing tags in the Mdb.
        /// </summary>
        /// <returns></returns>
        public bool Tags_SaveAllImagesToMDB()
        {
            return Tags_SaveAllImagesToMDB(DefaultSnapshot);
        }

        /// <summary>
        ///     Saves images for all existing tags in the Mdb.
        /// </summary>
        /// <param name="snap">The snapshot format the images will be created with.</param>
        /// <returns></returns>
        public bool Tags_SaveAllImagesToMDB(SprSnapShot snap)
        {
            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // Return null if the table retrieval failed
            if (tagTable == null) return false;

            // Return null if no tags exist
            if (tagTable.Rows.Count == 0) return false;

            for (int i = 0; i < tagTable.Rows.Count; i++)
            {
                var curTag = Tags_Get(Convert.ToInt32(tagTable.Rows[i]["tag_unique_id"]));

                if (!Tags_SaveImageToMDB(curTag.Id, snap, true))
                    return false;
            }

            return true;
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
            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // Set the name formatting
            nameFormat = nameFormat.Replace("##", "{0}");

            // Return null if the table retrieval failed
            if (tagTable == null) return false;

            // Return null if no tags exist
            if (tagTable.Rows.Count == 0) return false;

            for (int i = 0; i < tagTable.Rows.Count; i++)
            {
                var curTagNo = Convert.ToInt32(tagTable.Rows[i]["tag_unique_id"]);
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
            // Throw an expection if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Create the params
            var visValue = Convert.ToInt32(visible);

            // Create the view object
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            // Throw an exception if the DrViewDbl is null
            if (objViewdataDbl == null) throw SprExceptions.SprObjectCreateFail;

            // Set the view object as the SPR Application main view
            LastResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated annotation display
            objViewdataDbl.AllAnnotationsDisplay = visValue;

            // Update the global annotation visibility properties
            LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDisplay, visValue);
            LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoTextDisplay, visValue);
            LastResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDataDisplay, visValue);
                        
            // Update the main view in SPR
            LastResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        /// <summary>
        ///     Creates a new data field in the Mdb text_annotations table.
        ///     Returns true if the field already exists.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <returns>Indicates the success or failure of the table modification.</returns>
        public bool Annotations_AddDataField(string fieldName)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Add the tag field to the MDB database
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
            LastResult = DrApi.AnnotationCreateDbl(anno.Type, ref drAnno, out annoId);

            // Link the located object to the annotation
            LastResult = DrApi.AnnotationDataSet(annoId, anno.Type, ref drAnno, ref objId);

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
            LastResult = DrApi.ViewUpdate(1);
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
            LastResult = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Return null if the annotation locate failed
            if (annoId == 0) return null;

            // Set the annotation ID
            anno.Id = annoId;

            // Get the associated object ID
            int assocId;
            var drAnno = anno.DrAnnotationDbl;
            LastResult = DrApi.AnnotationDataGet(annoId, type, ref drAnno, out assocId);

            // Return null if the associated object id is zero
            if (assocId == 0) return null;

            // Set the assiciated object
            anno.AssociatedObject = GetObjectData(assocId);

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
            LastResult = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Return if the annotation locate was unsuccessful
            if (annoId == 0) return;

            // Delete the selected annotation
            LastResult = DrApi.AnnotationDelete(type, annoId, 0);

            // Update the main view
            LastResult = DrApi.ViewUpdate(1);
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
            LastResult = DrApi.AnnotationDeleteAll(type, 0);

            // Update the main view
            LastResult = DrApi.ViewUpdate(1);
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
            LastResult = DrApi.ViewUpdate(1);
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
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // Build the output image path (.BMP is forced before conversions)
            var imgPath = Path.Combine(outputDir, string.Format("{0}.bmp", imageName));

            // Get the current backface/endcap settings
            var orgBackfaces = GlobalOptionsGet(SprConstants.SprGlobalBackfacesDisplay);
            var orgEndcaps = GlobalOptionsGet(SprConstants.SprGlobalEndcapsDisplay);

            // Turn on view backfaces/endcaps as needed
            if (orgBackfaces == 0)
            GlobalOptionsSet(SprConstants.SprGlobalBackfacesDisplay, 1);
            if (orgEndcaps == 0)
            GlobalOptionsSet(SprConstants.SprGlobalEndcapsDisplay, 1);

            // Take the snapshot
            LastResult = DrApi.SnapShot(imgPath, snapShot.Flags, snapShot.DrSnapShot, 0);

            // Wait until SmartPlant Review is finished processing
            while (IsBusy) Thread.Sleep(100);

            // Reset the original settings if applicable
            if (orgBackfaces == 0)
                GlobalOptionsSet(SprConstants.SprGlobalBackfacesDisplay, 1);
            if (orgEndcaps == 0)
                GlobalOptionsSet(SprConstants.SprGlobalEndcapsDisplay, 1);

            // Return false if the snapshot doesn't exist
            if (!File.Exists(imgPath)) return null;

            // Format the snapshot if required
            if (snapShot.OutputFormat != SprSnapshotFormat.Bmp)
                imgPath = SprSnapShot.FormatSnapshot(imgPath, snapShot.OutputFormat);

            return imgPath;

        }
      
        /// <summary>
        ///     
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

            LastResult = DrApi.ExportPDF(path, quality, 1, 1, 1);
        }

        #endregion

        #region Views

        public void SetCenterPoint(double east, double north, double elevation)
        {
            SetCenterPoint(new SprPoint3D(east, north, elevation));
        }

        public void SetCenterPoint(SprPoint3D centerPoint)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Create the DrViewDbl
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            // Set the view object as the SPR Application main view
            LastResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated centerpoint
            objViewdataDbl.CenterUorPoint = centerPoint.DrPointDbl;

            // Update the main view in SPR
            LastResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
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
            LastResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated eyepoint
            objViewdataDbl.EyeUorPoint = eyePoint.DrPointDbl;

            // Update the main view in SPR
            LastResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        //public void GotoLocation

        #endregion

        #region Windows

        /// <summary>
        ///     Gets a SmartPlant Review application window.
        /// </summary>
        /// <param name="windowNo">The integer of the window to return.</param>
        /// <returns>The SprWindow containing the window properties.</returns>
        public SprWindow Window_Get(int windowNo)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Create the params
            int drHwnd;

            // Create the SPRWindow
            var curWin = new SprWindow();

            // Create the DrWindow
            dynamic objWin = Activator.CreateInstance(SprImportedTypes.DrWindow);
            LastResult = DrApi.WindowGet(windowNo, out objWin);

            // Get the window handle
            LastResult = DrApi.WindowHandleGet(windowNo, out drHwnd);

            // Set the window values
            if (objWin != null)
            {
                // Set the size
                curWin.Height = objWin.Height;
                curWin.Width = objWin.Width;

                // Set the position (0 if negative)
                curWin.Left = objWin.Left < 0 ? 0 : objWin.Left;
                curWin.Top = objWin.Top < 0 ? 0 : objWin.Top;

                // Set the handle
                curWin.WindowHandle = drHwnd;

                // Set the index
                curWin.Index = windowNo;

                // Return the window
                return curWin;
            }

            return null;
        }

        /// <summary>
        ///     Modifies a SmartPlant Review application window.
        /// </summary>
        /// <param name="window">The SprWindow containing the new properties.</param>
        public void Window_Set(SprWindow window)
        {
            if (!IsConnected)
                throw SprExceptions.SprNotConnected;

            // Create the DrWindow
            dynamic objWin = Activator.CreateInstance(SprImportedTypes.DrWindow);
            LastResult = DrApi.WindowGet(window.Index, out objWin);

            // Return if the DrWindow is null
            if (objWin == null) return;

            // Set the new size
            objWin.Height = window.Height;
            objWin.Width = window.Width;

            // Set the new position
            if (window.Left > 0) objWin.Left = window.Left;
            if (window.Top > 0) objWin.Top = window.Top;

            // Apply the updates
            LastResult = DrApi.WindowSet(window.Index, objWin);
        }

        #endregion
    }
}