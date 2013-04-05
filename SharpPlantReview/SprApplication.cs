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
        ///     The static Application used to set the class parent object
        /// </summary>
        internal static SprApplication ActiveApplication;

        private bool _disposed;

        /// <summary>
        ///     The collection of active running SprProcesses
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
            int sprResult = DrApi.SessionAttach(fileName);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidFileName:
                    throw SprExceptions.SprInvalidFileName;
                case SprConstants.SprErrorInvalidFileNumber:
                    throw SprExceptions.SprInvalidFileNumber;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorApiInternal:
                    throw SprExceptions.SprInternalError;
            }
        }

        /// <summary>
        ///     Exits the SmartPlant Review application.
        /// </summary>
        public void Exit()
        {
            // Throw an error if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;

            // Exit the SPR application
            int sprResult = DrApi.ExitViewer();

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }
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
            int sprResult = DrApi.PointLocateDbl(prompt, out abortFlag, ref returnPoint.DrPointDbl);

            // Return null if the locate operation was aborted
            if (abortFlag != 0) return null;

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }

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
            int sprResult = DrApi.PointLocateExtendedDbl(prompt, out abort, ref returnPoint.DrPointDbl,
                                                         ref targetPoint.DrPointDbl, out objId, flag);

            // Return null if the locate operation was aborted
            if (abort != 0) return null;

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }

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
            int sprResult = DrApi.ObjectLocateDbl(prompt, filterFlag, out returnId, ref refPoint.DrPointDbl);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }

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
            dynamic objData = returnData.DrObjectDataDbl;

            // Return null if the objectId is zero
            if (objectId == 0) return null;

            // Set the return object ID
            returnData.ObjectId = objectId;

            // Get the DataDbl object
            int sprResult = DrApi.ObjectDataGetDbl(objectId, 2, ref objData);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorInvalidObjectId:
                    throw SprExceptions.SprInvalidObjectId;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }

            // Iterate through the labels
            string lblName = string.Empty, lblValue = string.Empty;
            for (var i = 0; i < objData.LabelDataCount; i++)
            {
                // Get the label key/value pair
                sprResult = DrApi.ObjectDataLabelGet(ref lblName, ref lblValue, i);

                // Handle the errors (added a 0 break since this loops)
                switch (sprResult)
                {
                    case 0:
                        break;
                    case SprConstants.SprErrorInvalidParameter:
                        throw SprExceptions.SprInvalidParameter;
                    case SprConstants.SprErrorNoApi:
                        throw SprExceptions.SprNotConnected;
                }

                // Add the label information to the Labeldata dictionary
                try { returnData.LabelData.Add(lblName, lblValue); }
                catch (ArgumentException) { }
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
            int sprResult = DrApi.TextWindow(SprConstants.SprClearTextWindow, "Text View", string.Empty, 0);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
            }
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
            int sprResult = DrApi.TextWindow(SprConstants.SprClearTextWindow, titleText, mainText, 0);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
            }
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
            int sprResult = DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
                case SprConstants.SprErrorInvalidParameter:
                    if (orgTitle == null || orgText == null) break;  // Ignore null parameters
                    throw SprExceptions.SprInvalidParameter;
            }

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
            int sprResult = DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
                case SprConstants.SprErrorInvalidParameter:
                    if (orgTitle == null || orgText == null) break; // Ignore null parameters
                    throw SprExceptions.SprInvalidParameter;
            }

            // Set an empty string for null values
            return orgText ?? (string.Empty);

            // Return the text
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
            return DbMethods.AddDbField(MdbPath, fieldName);
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
            int sprResult = DrApi.TagDelete(tagNo, 0);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorInvalidTag:
                    throw SprExceptions.SprInvalidTag;
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }

            // Set the deleted tag as the next tag number
            if (setAsNextTag)
                Tags_SetNextTag(tagNo);

            // Update the SmartPlant Review main view
            DrApi.ViewUpdate(1);
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
            TextWindow_Update(curTag.TagData["tag_text"].ToString(), string.Format("Tag {0}", tagNo));

            // Locate the desired tag on the main screen with the specified visibility
            int sprResult = DrApi.GotoTag(tagNo, 0, Convert.ToInt32(displayTag));

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorInvalidTag:
                    throw SprExceptions.SprInvalidTag;
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
            }
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
                returnTag.TagData[col.ColumnName] = rowFilter[0][col];
            }

            // Set the tag as placed
            returnTag.IsPlaced = true;

            // Return the tag
            return returnTag;
        }

        /// <summary>
        ///     Retrieves all the existing tags in the Mdb tag_data table.
        /// </summary>
        /// <returns>List of SprTags retrieved.</returns>
        public List<SprTag> Tags_GetAll()
        {
            throw new NotImplementedException();
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

            // Create the params
            dynamic tagKey = Activator.CreateInstance(SprImportedTypes.DrKey);

            // Create the origin point
            var tagOrigin = new SprPoint3D();

            // Get an object on screen and set the origin point to its location
            var objId = -1;
            objId = GetObjectId("SELECT TAG START POINT", ref tagOrigin);

            // Exit if the object selection failed
            if (objId == -1) return;

            // Get the tag leader point using the origin for depth
            var tagLeader = GetPoint("SELECT TAG LEADER LOCATION", tagOrigin);

            // Exit if the leader point is not set
            if (tagLeader == null) return;

            // Set the tag registry values
            SprUtilities.SetTagRegistry(tag);

            // Place the tag
            int sprResult = DrApi.TagSetDbl(tag.TagNumber, 0, tag.Flags, ref tagLeader.DrPointDbl,
                                            ref tagOrigin.DrPointDbl, tagKey, tag.Text);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorOutOfMemory:
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidTag:
                    throw SprExceptions.SprInvalidTag;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorTagExists:
                    throw SprExceptions.SprTagExists;
            }

            // Clear the tag registry
            SprUtilities.ClearTagRegistry();

            // Update the text window
            TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.TagNumber));

            // Reference the placed tag
            tag = Tags_Get(tag.TagNumber);
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
            var objId = -1;
            objId = GetObjectId("SELECT NEW TAG START POINT", ref tagOrigin);

            // Get the tag leader point
            var tagLeader = GetPoint("SELECT NEW LEADER LOCATION");

            // Throw an exception if either of the point retrievals failed
            if (objId == -1 || tagLeader == null) throw SprExceptions.SprNullPoint;

            // Create DrKey object
            dynamic tagKey = Activator.CreateInstance(SprImportedTypes.DrKey);

            // Throw an exception if the key is null
            if (tagKey == null) throw SprExceptions.SprObjectCreateFail;

            // Set the edit flag on the existing tag
            tag.Flags |= SprConstants.SprTagEdit;

            // Update the tag with the new leader points
            int sprResult = DrApi.TagSetDbl(tag.TagNumber, 0, tag.Flags, tagLeader.DrPointDbl,
                                                tagOrigin.DrPointDbl, tagKey, tagText);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorOutOfMemory:
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidTag:
                    throw SprExceptions.SprInvalidTag;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorTagExists:
                    throw SprExceptions.SprTagExists;
            }

            // Reference the placed tag
            tag = Tags_Get(tag.TagNumber);

            // Flip the tag 180 degrees.  Intergraph is AWESOME!
            var newOrigin = tag.LeaderPoint;
            var newLeader = tag.OriginPoint;
            tag.LeaderPoint = newLeader;
            tag.OriginPoint = newOrigin;

            // Queue the Mdb update in the threadpool
            ThreadPool.QueueUserWorkItem(Tags_Update, tag);

            // Update the text window
            TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.TagNumber));

            // Update the main view
            DrApi.ViewUpdate(1);
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
            var rowFilter = string.Format("tag_unique_id = {0}", tag.TagNumber);
            var tblFilter = tagTable.Select(rowFilter);

            // Iterate through each dictionary key/value pair
            foreach (var kvp in tag.TagData)
            
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
            var rowFilter = string.Format("tag_unique_id = {0}", tag.TagNumber);
            var tblFilter = tagTable.Select(rowFilter);

            // Iterate through each dictionary key/value pair
            foreach (var kvp in tag.TagData)

                // Set the values for the selected tag
                tblFilter[0][kvp.Key] = kvp.Value;

            // Push the the updated table
            DbMethods.UpdateDbTable(MdbPath, rowFilter, tagTable);
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
            int sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidView:
                    throw SprExceptions.SprInvalidView;
            }

            // Apply the updated annotation display
            objViewdataDbl.AllAnnotationsDisplay = visValue;

            // Update the global annotation visibility properties
            sprResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDisplay, visValue);
            sprResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoTextDisplay, visValue);
            sprResult = DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDataDisplay, visValue);
            
            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidGlobal:
                    throw SprExceptions.SprInvalidGlobal;
            }
            
            // Update the main view in SPR
            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidView:
                    throw SprExceptions.SprInvalidView;
            }
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
            int assocId = GetObjectId("SELECT A POINT ON AN OBJECT TO LOCATE THE ANNOTATION", ref leaderPoint);

            // Exit if the leaderObj was set
            if (assocId == -1) return;           

            // Get the annotation center point using the leaderpoint for depth calculation
            var centerPoint = GetPoint("SELECT THE CENTER POINT FOR THE ANNOTATION LABEL", leaderPoint);

            // Exit if the centerObj was set
            if (centerPoint == null) return;

            // Set the annotation points
            anno.LeaderPoint = leaderPoint;
            anno.CenterPoint = centerPoint;

            // Place the annotation on screen
            int annoId;
            int sprResult = DrApi.AnnotationCreateDbl(anno.Type, ref anno.DrAnnotationDbl, out annoId);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
            }

            // Set the annotation Id
            anno.AnnotationId = annoId;

            // Link the located object to the annotation
            sprResult = DrApi.AnnotationDataSet(annoId, anno.Type, ref anno.DrAnnotationDbl, ref assocId);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
            }

            // Get the associated object data
            anno.AssociatedObject = GetObjectData(assocId);
            anno.AssociatedObject.SelectedPoint = anno.LeaderPoint;

            // Update the main view
            DrApi.ViewUpdate(1);
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
            var returnAnno = new SprAnnotation();

            // Set the SPR application visible
            Activate();

            // Prompt the user to select the annotation
            var msg = string.Format("SELECT THE DESIRED {0} ANNOTATION", type.ToUpper());
            int sprResult = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
            }

            // Return null if the annotation locate failed
            if (annoId == 0) return null;

            // Set the annotation ID
            returnAnno.AnnotationId = annoId;

            // Get the associated object ID
            int assocId;
            sprResult = DrApi.AnnotationDataGet(annoId, type, ref returnAnno.DrAnnotationDbl, out assocId);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
                case SprConstants.SprErrorApiOutOfMemory:
                    throw SprExceptions.SprOutOfMemory;
            }

            // Return null if the associated object id is zero
            if (assocId == 0) return null;

            // Set the assiciated object
            returnAnno.AssociatedObject = GetObjectData(assocId);

            // Return the completed annotation
            return returnAnno;
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
            int sprResult = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
            }

            // Return if the annotation locate was unsuccessful
            if (annoId == 0) return;

            // Delete the selected annotation
            sprResult = DrApi.AnnotationDelete(type, annoId, 0);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
                case SprConstants.SprErrorInvalidObjectId:
                    throw SprExceptions.SprInvalidObjectId;
            }

            // Update the main view
            DrApi.ViewUpdate(1);
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
            int sprResult = DrApi.AnnotationDeleteAll(type, 0);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidParameter:
                    throw SprExceptions.SprInvalidParameter;
                case SprConstants.SprErrorInvalidAnnoType:
                    throw SprExceptions.SprInvalidAnnoType;
                case SprConstants.SprErrorInvalidObjectId:
                    throw SprExceptions.SprInvalidObjectId;
            }

            // Update the main view
            DrApi.ViewUpdate(1);
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
            DrApi.ViewUpdate(1);
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
        public bool TakeSnapshot(SprSnapShot snapShot, string imageName, string outputDir)
        {
            // Throw an exception if not connected
            if (!IsConnected) throw SprExceptions.SprNotConnected;
            
            // Build the output image path (.BMP is forced before conversions)
            var imgPath = Path.Combine(outputDir, string.Format("{0}.bmp", imageName));

            // Take the snapshot
            int sprResult = DrApi.SnapShot(imgPath, snapShot.Flags, snapShot.DrSnapShot, 0);

            // Handle the errors
            switch (sprResult)
            {
                case SprConstants.SprErrorNoApi:
                    throw SprExceptions.SprNotConnected;
                case SprConstants.SprErrorInvalidFileName:
                    throw SprExceptions.SprInvalidFileName;
                case SprConstants.SprErrorDirectoryWriteFailure:
                    throw SprExceptions.SprDirectoryWriteFailure;
                case SprConstants.SprErrorFileWriteFailure:
                    throw SprExceptions.SprFileWriteFailure;
                case SprConstants.SprErrorFileExists:
                    throw SprExceptions.SprFileExists;
            }

            // Return false if the snapshot doesn't exist
            if (!File.Exists(imgPath)) return false;

            // Format the snapshot if required
            return snapShot.OutputFormat == SprSnapshotFormat.Bmp || 
                SprSnapShot.FormatSnapshot(imgPath, snapShot.OutputFormat);
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
            DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated centerpoint
            objViewdataDbl.CenterUorPoint = centerPoint.DrPointDbl;

            // Update the main view in SPR
            DrApi.ViewSetDbl(0, ref objViewdataDbl);
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
            DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated eyepoint
            objViewdataDbl.EyeUorPoint = eyePoint.DrPointDbl;

            // Update the main view in SPR
            DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

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
            DrApi.WindowGet(windowNo, out objWin);

            // Get the window handle
            DrApi.WindowHandleGet(windowNo, out drHwnd);

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
            DrApi.WindowGet(window.Index, out objWin);

            // Return if the DrWindow is null
            if (objWin == null) return;

            // Set the new size
            objWin.Height = window.Height;
            objWin.Width = window.Width;

            // Set the new position
            if (window.Left > 0) objWin.Left = window.Left;
            if (window.Top > 0) objWin.Top = window.Top;

            // Apply the updates
            DrApi.WindowSet(window.Index, objWin);
        }

        #endregion
    }
}