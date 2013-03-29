using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Provides methods and properties for interacting with SmartPlant Review.
    /// </summary>
    public class Application
    {
        #region Application Properties

        /// <summary>
        ///     The active COM reference to the DrAPi class.
        /// </summary>
        internal dynamic DrApi;

        /// <summary>
        ///     Determines if an active connection to SmartPlant Review is established.
        /// </summary>
        public bool IsConnected
        {
            get { return (DrApi != null); }
        }

        /// <summary>
        ///     Gets the MDB path to the active review session.
        /// </summary>
        public string MdbPath
        {
            get
            {
                if (!IsConnected)
                    return null;
                // params for holding the return data
                var dirPath = string.Empty;
                var mdbName = string.Empty;

                // Set the global option to retrieve database file info
                DrApi.GlobalOptionsSet(Constants.G_API_FILE_INFO_MODE, 1);

                // Get the directory path
                DrApi.FilePathFromNumber(1, ref dirPath);

                // Get the MDB FileName
                DrApi.FileNameFromNumber(1, ref mdbName);

                // Reset the global variable
                DrApi.GlobalOptionsSet(Constants.G_API_FILE_INFO_MODE, 0);

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
                if (!IsConnected) return null;

                // Create the return parameter
                string vueName = null;

                // Set the global variable
                DrApi.GlobalOptionsSet(Constants.G_API_FILE_INFO_MODE, 1);

                // Get the VUE file name
                DrApi.FileNameFromNumber(0, ref vueName);

                // Reset the global variable
                DrApi.GlobalOptionsSet(Constants.G_API_FILE_INFO_MODE, 0);

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
                if (!IsConnected)
                    return -1;
                // Create the parameter
                var returnAnno = -1;

                // Retrieve the MDB site table
                var siteTable = DbMethods.GetDbTable(MdbPath, "site_table");

                if (siteTable != null)

                    // Set the next annotation number
                    returnAnno = (int) siteTable.Rows[0]["next_text_anno_id"];

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
                if (!IsConnected)
                    return false;
                // Get the SPR Application process
                uint procId;
                DrApi.ProcessIdGet(out procId);

                // Get a starting point to measure
                var sw = new Stopwatch();

                // Wait until the application is idle (10 ms max)
                sw.Start();
                Process.GetProcessById((int) procId).WaitForInputIdle(10);
                sw.Stop();

                // Return true if the application waited the full test period
                return (sw.ElapsedMilliseconds >= 10);
            }
        }

        /// <summary>
        ///     Gets the process ID of the SmartPlant Review application.
        /// </summary>
        public uint ProcessId
        {
            get
            {
                if (!IsConnected)
                    return 0;
                uint procId;
                DrApi.ProcessIdGet(out procId);

                return procId;
            }
        }

        /// <summary>
        ///     The primary SmartPlant Review application window.
        /// </summary>
        public Window ApplicationWindow
        {
            get { return Window_Get(Constants.SMARTPLANT_REVIEW_WIN); }
            set { Window_Set(value); }
        }

        /// <summary>
        ///     The Main view window inside the SmartPlant Review application.
        /// </summary>
        public Window MainWindow
        {
            get { return Window_Get(Constants.MAIN_WIN); }
            set { Window_Set(value); }
        }

        /// <summary>
        ///     The Plan view window inside the SmartPlant Review application.
        /// </summary>
        public Window PlanWindow
        {
            get { return Window_Get(Constants.PLAN_WIN); }
            set { Window_Set(value); }
        }

        /// <summary>
        ///     The Elevation window inside the SmartPlant Review application.
        /// </summary>
        public Window ElevationWindow
        {
            get { return Window_Get(Constants.ELEV_WIN); }
            set { Window_Set(value); }
        }

        /// <summary>
        ///     The Text window inside the SmartPlant Review application.
        /// </summary>
        public Window TextWindow
        {
            get { return Window_Get(Constants.TEXT_WIN); }
            set { Window_Set(value); }
        }

        #endregion

        // Application initializer
        public Application()
        {
            // Set the static application for class parent referencing
            SmartPlantReview.ActiveApplication = this;

            // Get the application windows
        }

        #region General

        public void Activate()
        {
            NativeWin32.SetForegroundWindow(ApplicationWindow.WindowHandle);
        }

        public bool Connect()
        {
            // Clear the current Application
            DrApi = null;

            // Check for running instances of SPR
            var apps = Process.GetProcessesByName("spr");

            // If at least one instance of SPR is running
            if (apps.Length > 0)
            {
                // Get the running instance of SmartPlant Review
                DrApi = Activator.CreateInstance(ImportedTypes.DrApi);
            }

            // Return the connection state
            return IsConnected;
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(DrApi);
            DrApi = null;
        }

        public bool Open(string fileName)
        {
            if (IsConnected)
            {
                int sprResult = DrApi.SessionAttach(fileName);
                return (sprResult == 0);
            }
            throw Exceptions.ApiNotConnected;
        }

        public void Exit()
        {
            if (IsConnected)
            {
                // Exit the SPR application
                DrApi.ExitViewer();
            }
            else throw Exceptions.ApiNotConnected;
        }

        public void RefreshData()
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;
            // Set SPR to the front
            Activate();

            // Send the update command
            SendKeys.SendWait("%TR");
        }

        #endregion

        #region Point Select

        public Point3D GetPoint(string prompt)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;
            // Create the DrPointDbl
            dynamic objPoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);

            // Create the params
            int abortFlag;

            // Set the SPR application visible
            Activate();

            // Prompt the user for a 3D point inside SPR
            int sprResult = DrApi.PointLocateDbl(prompt, out abortFlag, ref objPoint);

            // On error return null
            return sprResult != 0 ? null : new Point3D(objPoint.East, objPoint.North, objPoint.Elevation);

            // Return the new point
        }

        public int GetObjectId(string prompt)
        {
            var pt = new Point3D();
            return GetObjectId(prompt, ref pt);
        }

        public int GetObjectId(string prompt, ref Point3D refPoint)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;
            // Create the DrPointDbl
            dynamic objPoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);

            // Create the params
            const int filterFlag = 0;
            int returnId;

            // Set the SPR application visible
            Activate();

            // Prompt the user for a 3D point inside SPR
            int sprResult = DrApi.ObjectLocateDbl(prompt, filterFlag, out returnId, ref objPoint);

            // Link the objPoint to the reference point
            refPoint.East = objPoint.East;
            refPoint.North = objPoint.North;
            refPoint.Elevation = objPoint.Elevation;

            // Return the object ID
            if (sprResult == 0)
                return returnId;
            return -1;
        }

        public ObjectData GetObjectData(int objectId)
        {
            return GetObjectData(string.Empty, false, objectId);
        }

        public ObjectData GetObjectData(string prompt)
        {
            return GetObjectData(prompt, false, 0);
        }

        public ObjectData GetObjectData(string prompt, bool singleObject)
        {
            return GetObjectData(prompt, singleObject, 0);
        }

        public ObjectData GetObjectData(string prompt, bool singleObject, int objectId)
        {
            if (IsConnected)
            {
                // Create the return object
                var returnData = new ObjectData();
                dynamic objData = returnData.DrObjectDataDbl;

                // Create the DrPointDbl
                dynamic selectPoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);

                // Create the params
                int flags = 0;
                if (singleObject) flags += 1;
                int sprResult = 0;

                // If the object ID was not passed
                if (objectId == 0)
                {
                    // Set the SPR application visible
                    Activate();

                    // Get the object ID
                    sprResult = DrApi.ObjectLocateDbl(prompt, flags, out objectId, ref selectPoint);

                    // Set the selected point
                    returnData.SelectedPoint = new Point3D(selectPoint.East, selectPoint.North, selectPoint.Elevation);
                }

                // Set the return object ID
                returnData.ObjectId = objectId;

                // If the ObjectLocateDbl method suceeded
                if (sprResult == 0 && objectId != 0)
                {
                    // Get the DataDbl object
                    sprResult = DrApi.ObjectDataGetDbl(objectId, 2, ref objData);

                    // If the ObjectDataGetDb method suceeded
                    if (sprResult == 0)
                    {
                        // Iterate through the labels
                        string lblName = string.Empty, lblValue = string.Empty;
                        for (int i = 0; i < objData.LabelDataCount; i++)
                        {
                            // Get the label key/value pair
                            sprResult = DrApi.ObjectDataLabelGet(ref lblName, ref lblValue, i);
                            if (sprResult == 0)
                            {
                                try
                                {
                                    // Add the current label to the dictionary
                                    returnData.LabelData.Add(lblName, lblValue);
                                }
                                catch (ArgumentException)
                                {
                                }
                            }
                        }

                        // Return the data object
                        return returnData;
                    }
                }

                // Return null on error
                return null;
            }
            throw Exceptions.ApiNotConnected;
        }

        public ObjectData GetObjectData(string prompt, bool singleObject, Point3D target)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;
            // Create the return object
            var returnData = new ObjectData();

            // Create the DrPointDbl points
            dynamic selectPoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);
            dynamic targetPoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);

            // Set the target point values
            targetPoint.East = target.East;
            targetPoint.North = target.North;
            targetPoint.Elevation = target.Elevation;

            // Create the params
            var flags = 0;
            if (singleObject) flags += 2;
            int objId;
            int abort;

            // Set the SPR application visible
            Activate();

            // Get the object ID and selected point
            int sprResult = DrApi.PointLocateExtendedDbl(prompt, out abort, ref selectPoint,
                                                         ref targetPoint, out objId, flags);

            // Set the selected point
            returnData.SelectedPoint = new Point3D(selectPoint.East, selectPoint.North, selectPoint.Elevation);

            // If the PointLocateExtendedDbl method suceeded
            return sprResult == 0 ? returnData : null;

            // Return null on error
        }

        #endregion

        #region Text Window

        public void TextWindow_Clear()
        {
            if (IsConnected)
            {
                // Send a blank string to the application text window
                DrApi.TextWindow(Constants.TEXT_WIN_CLR, "Text View", string.Empty, 0);
            }
            else throw Exceptions.ApiNotConnected;
        }

        public void TextWindow_Update(string mainText)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Get the existing title
            string existTitle = TextWindow_GetTitle();

            // Set the text window without changing the title
            TextWindow_Update(mainText, existTitle);
        }

        public void TextWindow_Update(string mainText, string titleText)
        {
            if (IsConnected)
            {
                // Set the text window and title contents
                DrApi.TextWindow(Constants.TEXT_WIN_CLR, titleText, mainText, 0);
            }
            else throw Exceptions.ApiNotConnected;
        }

        public string TextWindow_GetTitle()
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Params for retrieving SPR data
            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            // Get the existing text window values
            DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Return the title
            return orgTitle;
        }

        public string TextWindow_GetText()
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Params for retrieving SPR data
            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            // Get the existing text window values
            DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Return the text
            return orgText;
        }

        #endregion

        #region MDB Database

        public bool MDB_AddTagDataField(string fieldName)
        {
            if (IsConnected)
            {
                // Add the tag field to the MDB database
                return DbMethods.AddDbField(MdbPath, fieldName);
            }
            throw Exceptions.ApiNotConnected;
        }

        #endregion

        #region Tagging

        public bool Tags_Add(Tag tag)
        {
            return false;
        }

        public bool Tags_Delete(int tagNo)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Delete the desired tag
            int sprResult = DrApi.TagDelete(tagNo, 0);
            if (sprResult == 0)
            {
                // Clear the text window
                TextWindow_Clear();

                // Fix the next tag in the MDB database
                return Tags_FixNextTag();
            }

            return false;
        }

        public void Tags_Display(TagVisibility displayState)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Set SPR to the front
            Activate();

            // Get the menu alias character from the enumerator
            var alias = Char.ConvertFromUtf32((int) displayState);

            // Set the tag visibility
            SendKeys.SendWait(string.Format("%GS{0}", alias));
        }

        public bool Tags_FixNextTag()
        {
            if (IsConnected)
            {
                // Get the tags
                var tagTable = DbMethods.GetDbTable(MdbPath, "tag_table");

                // Retrieve the site table
                var siteTable = DbMethods.GetDbTable(MdbPath, "site_table");

                // If tags exist
                if (tagTable.Rows.Count > 0)

                    // Set the next tag to the highest tag value + 1
                    siteTable.Rows[0]["next_tag_id"] =
                        Convert.ToInt32(tagTable.Rows[tagTable.Rows.Count - 1]["tag_unique_id"]) + 1;
                else

                    // Set the next tag to 1
                    siteTable.Rows[0]["next_tag_id"] = 1;

                // Return the result of the table update
                return DbMethods.UpdateDbTable(MdbPath, siteTable);
            }
            throw Exceptions.ApiNotConnected;
        }

        public void Tags_Goto(int tagNo)
        {
            Tags_Goto(tagNo, true);
        }

        public void Tags_Goto(int tagNo, bool displayText)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            if (displayText)
            {
                // Get the tag data
                var curTag = Tags_Get(tagNo);

                // Update the text window with the tag information
                TextWindow_Update(curTag.TagData["tag_text"].ToString(), string.Format("Tag {0}", tagNo));
            }
            else
            {
                // Clear the text window
                TextWindow_Clear();
            }

            // Show the desired tag on the main screen
            DrApi.GotoTag(tagNo, 0, 1);
        }

        public Tag Tags_Get(int tagNo)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the new tag
            var returnTag = new Tag();

            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // If tags exist
            if (tagTable.Rows.Count > 0)
            {
                // Create the row filter for the desired tag
                var rowFilter = tagTable.Select(string.Format("tag_unique_id = '{0}'", tagNo));

                // Iterate through each column
                foreach (DataColumn col in tagTable.Columns)
                {
                    // Add the key/value from the first filtered row to the dictionary
                    returnTag.TagData[col.ColumnName] = rowFilter[0][col];
                }

                // Return the tag
                return returnTag;
            }

            // Return nothing
            return null;
        }

        public List<Tag> Tags_GetAll()
        {
            return null;
        }

        public bool Tags_Place(string tagText)
        {
            var tag = new Tag {Text = tagText};

            return Tags_Place(ref tag);
        }

        public bool Tags_Place(ref Tag tag)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            dynamic tagKey = Activator.CreateInstance(ImportedTypes.DrKey);
            const int tagFlag = Constants.TAG_LEADER;

            // Set the tag key
            tagKey.LabelKey1 = 1;

            // Get the tagged object location
            var tagOrigin = GetPoint("SELECT TAG START POINT");

            // Get the tag leader point
            var tagLeader = GetPoint("SELECT TAG LEADER LOCATION");

            // Place the tag
            int sprResult = DrApi.TagSetDbl(tag.TagNumber, 0, tagFlag, ref tagLeader, ref tagOrigin, tagKey, tag.Text);

            if (sprResult == 0)
            {
                // Update the text window
                TextWindow_Update(tag.Text, string.Format("Tag {0}", tag.TagNumber));

                // Update the complete tag data
                return Tags_Update(tag);
            }

            return false;
        }

        public bool Tags_PlaceLeader(int tagNo)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            var existData = string.Empty;

            // Get the tagged object location
            var tagOrigin = GetPoint("SELECT NEW TAG START POINT");

            // Get the tag leader point
            var tagLeader = GetPoint("SELECT NEW LEADER LOCATION");

            // If the point collection was successful
            if (tagOrigin != null && tagLeader != null)
            {
                // Get the old tag Data
                int dataLength;
                DrApi.TagDataGet(tagNo, 0, out dataLength, ref existData);

                // Delete the old tag
                int sprResult = DrApi.TagDelete(tagNo, 0);

                // Update the text window
                TextWindow_Update("Updating tag leader..., ", string.Format("Tag {0}", tagNo));

                // If no errors exist
                if (sprResult == 0)
                {
                    // Update the next tag value
                    if (Tags_SetNextTag(tagNo))
                    {
                        // Create the params
                        dynamic tagKey = Activator.CreateInstance(ImportedTypes.DrKey);
                        const int tagFlag = Constants.TAG_LEADER;

                        // Set the tag key
                        tagKey.LabelKey1 = 1;

                        // Create the new tag with the updated leader
                        sprResult = DrApi.TagSetDbl(tagNo, 0, tagFlag, tagLeader, tagOrigin, tagKey, existData);

                        // Update the text window
                        TextWindow_Update(existData, "Tag " + tagNo);

                        // Return the error status
                        return (sprResult == 0);
                    }
                }
            }

            return false;
        }

        public bool Tags_SetNextTag(int tagNo)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Get the current database
            var siteTable = DbMethods.GetDbTable(MdbPath, "site_table");

            // If the table was collected
            if (siteTable != null)
            {
                // Get the top row
                var row = siteTable.Rows[0];

                // Set the next tag value
                row["next_tag_id"] = tagNo;
                return DbMethods.UpdateDbTable(MdbPath, siteTable);
            }

            return false;
        }

        public bool Tags_Update(Tag tag)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Retrieve the site table
            var tagTable = DbMethods.GetDbTable(MdbPath, "tag_data");

            // If tags exist
            if (tagTable.Rows.Count > 0)
            {
                // Create the row filter for the specified tag
                var rowFilter = string.Format("tag_unique_id = {0}", tag.TagNumber);

                // Iterate through each dictionary key/value pair
                foreach (var kvp in tag.TagData)
                {
                    // Set the values for the selected tag
                    tagTable.Rows[tag.TagNumber - 1][kvp.Key] = kvp.Value;
                }

                // Return the result of the table update
                return DbMethods.UpdateDbTable(MdbPath, rowFilter, tagTable);
            }

            return false;
        }

        #endregion

        #region Annotation

        public void Annotations_Display(bool visible)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            int sprResult;

            var visValue = Convert.ToInt32(visible);

            // Create the view object
            dynamic objViewdataDbl = Activator.CreateInstance(ImportedTypes.DrViewDbl);

            // Set the view object as the SPR Application main view
            sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated annotation display
            objViewdataDbl.AllAnnotationsDisplay = visValue;

            // Update the global properties
            DrApi.GlobalOptionsSet(Constants.G_ANNO_DISPLAYED, visValue);
            DrApi.GlobalOptionsSet(Constants.G_ANNO_TEXT_DISPLAYED, visValue);

            // Update the main view in SPR
            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        public void Annotations_Place(ref Annotation annotation)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            int annoId;
            int sprResult;

            // Get the annotation leader point
            var leaderObj = GetObjectData("SELECT A POINT ON AN OBJECT TO LOCATE THE ANNOTATION");

            // Check if the leaderObj was set
            if (leaderObj == null) return;

            // Set the annotation values
            annotation.LeaderPoint = leaderObj.SelectedPoint;
            var assocId = leaderObj.ObjectId;

            // Get the annotation center point
            var centerObj = GetObjectData("SELECT THE CENTER POINT FOR THE ANNOTATION LABEL", true,
                                          annotation.LeaderPoint);

            // Check if the centerObj was set
            if (centerObj == null) return;

            // Set the annotation CenterPoint
            annotation.CenterPoint = centerObj.SelectedPoint;

            // Place the annotation on screen
            sprResult = DrApi.AnnotationCreateDbl(annotation.Type, ref annotation.DrAnnotationDbl, out annoId);
            annotation.AnnotationId = annoId;

            // Link the located object to the annotation
            sprResult = DrApi.AnnotationDataSet(annoId, annotation.Type, ref annotation.DrAnnotationDbl, ref assocId);
            annotation.AssociatedObject = GetObjectData(assocId);
            annotation.AssociatedObject.SelectedPoint = annotation.LeaderPoint;

            // Update the main view
            DrApi.ViewUpdate(1);
        }

        public void Annotations_Add(Annotation annotation)
        {
        }

        public void Annotations_Add(List<Annotation> annotations)
        {
            foreach (var anno in annotations)
                Annotations_Add(anno);
        }

        public Annotation Annotations_Select(string type)
        {
            if (IsConnected)
            {
                // Create the params
                int sprResult;
                int annoId;

                // Create the return annotation object
                var returnAnno = new Annotation();

                // Set the SPR application visible
                Activate();

                // Prompt the user to select the annotation
                var msg = string.Format("SELECT THE DESIRED {0} ANNOTATION", type.ToUpper());
                sprResult = DrApi.AnnotationLocate(type, msg, 0, out annoId);

                // If the annotation locate was successful
                if (sprResult == 0 && annoId != 0)
                {
                    // Set the annotation ID
                    returnAnno.AnnotationId = annoId;

                    // Get the associated object ID
                    int assocId;
                    sprResult = DrApi.AnnotationDataGet(annoId, type, ref returnAnno.DrAnnotationDbl, out assocId);

                    // If the associated object was retrieved successfully
                    if (sprResult == 0 && assocId != 0)
                    {
                        // Set the assiciated object
                        returnAnno.AssociatedObject = GetObjectData(assocId);

                        // Return the finished annotation
                        return returnAnno;
                    }
                }

                // Return null on error
                return null;
            }
            throw Exceptions.ApiNotConnected;
        }

        public void Annotations_Delete(string type)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            int sprResult;
            int annoId;

            // Create the return annotation object
            var returnAnno = new Annotation();

            // Set the SPR application visible
            Activate();

            // Prompt the user to select the annotation
            var msg = string.Format("SELECT THE {0} ANNOTATION TO DELETE", type.ToUpper());
            sprResult = DrApi.AnnotationLocate(type, msg, 0, out annoId);

            // Return if the annotation locate was unsuccessful
            if (sprResult != 0 || annoId == 0) return;

            // Delete the selected annotation
            sprResult = DrApi.AnnotationDelete(type, annoId, 0);

            // Update the main view
            DrApi.ViewUpdate(Constants.MAIN_VIEW);
        }

        public void Annotations_DeleteType(string type)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Delete all annotations matching the provided type
            int sprResult = DrApi.AnnotationDeleteAll(type, 0);

            // Update the main view
            DrApi.ViewUpdate(1);
        }

        public void Annotations_DeleteAll()
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Set the params
            int sprResult;

            // Get the annotation types
            var typeTable = DbMethods.GetDbTable(MdbPath, "text_annotation_types");

            // If types exist
            if (typeTable.Rows.Count <= 0) return;

            // Iterate through each annotation type
            for (var i = typeTable.Rows.Count - 1; i >= 0; i--)
            {
                // Delete all annotations matching the current type
                sprResult = DrApi.AnnotationDeleteAll(typeTable.Rows[i]["name"].ToString(), 0);
            }

            // Update the main view
            DrApi.ViewUpdate(1);
        }

        #endregion

        #region Snapshot

        public bool TakeSnapshot(SnapShot snapShot, string imageName, string outputDir)
        {
            if (IsConnected)
            {
                // Build the output image path (.BMP is forced before conversions)
                var imgPath = Path.Combine(outputDir, string.Format("{0}.bmp", imageName));

                // Take the snapshot
                int sprResult = DrApi.SnapShot(imgPath, snapShot.Flags, snapShot.DrSnapShot, 0);

                // Check if the file exists
                if (!File.Exists(imgPath))
                    return false;

                // Format the snapshot if required
                return snapShot.OutputFormat == SnapshotFormat.Bmp || SnapShot.FormatSnapshot(imgPath, snapShot.OutputFormat);
            }
            throw Exceptions.ApiNotConnected;
        }

        #endregion

        #region Views

        public void SetCenterPoint(double east, double north, double elevation)
        {
            SetCenterPoint(new Point3D(east, north, elevation));
        }

        public void SetCenterPoint(Point3D centerPoint)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            int sprResult;

            // Create the DrViewDbl
            dynamic objViewdataDbl = Activator.CreateInstance(ImportedTypes.DrViewDbl);

            // Create the DrPointDbl
            dynamic objCenterPoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);

            // Set the centerpoint values
            objCenterPoint.East = centerPoint.East;
            objCenterPoint.North = centerPoint.North;
            objCenterPoint.Elevation = centerPoint.Elevation;

            // Set the view object as the SPR Application main view
            sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated centerpoint
            objViewdataDbl.CenterUorPoint = objCenterPoint;

            // Update the main view in SPR
            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        public void SetEyePoint(double east, double north, double elevation)
        {
            SetEyePoint(new Point3D(east, north, elevation));
        }

        public void SetEyePoint(Point3D eyePoint)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            int sprResult;

            // Create the DrViewDbl
            dynamic objViewdataDbl = Activator.CreateInstance(ImportedTypes.DrViewDbl);

            // Create the DrPointDbl
            dynamic objEyePoint = Activator.CreateInstance(ImportedTypes.DrPointDbl);

            // Set the centerpoint values
            objEyePoint.East = eyePoint.East;
            objEyePoint.North = eyePoint.North;

            // Set the view object as the SPR Application main view
            sprResult = DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated eyepoint
            objViewdataDbl.EyeUorPoint = objEyePoint;

            // Update the main view in SPR
            sprResult = DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        #endregion

        #region Windows

        public Window Window_Get(int windowNo)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the params
            int drHwnd;

            // Create the SPRWindow
            var curWin = new Window();

            // Create the DrWindow
            dynamic objWin = Activator.CreateInstance(ImportedTypes.DrWindow);
            DrApi.WindowGet(windowNo, out objWin);

            // Get the window handle
            DrApi.WindowHandleGet(windowNo, out drHwnd);

            // Set the window values
            if (objWin != null)
            {
                // Set the size
                curWin.Height = objWin.Height;
                curWin.Width = objWin.Width;

                // Set the position
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

        public void Window_Set(Window window)
        {
            if (!IsConnected)
                throw Exceptions.ApiNotConnected;

            // Create the DrWindow
            dynamic objWin = Activator.CreateInstance(ImportedTypes.DrWindow);
            DrApi.WindowGet(window.Index, out objWin);

            // Set the DrWindow values
            if (objWin != null)
            {
                // Set the new size
                objWin.Height = window.Height;
                objWin.Width = window.Width;

                // Set the new position
                if (window.Left > 0) objWin.Left = window.Left;
                if (window.Top > 0) objWin.Top = window.Top;

                // Apply the updates
                DrApi.WindowSet(window.Index, objWin);
            }
        }

        #endregion
    }
}