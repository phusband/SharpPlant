//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the properties for creating a tag in SmartPlant Review.
    /// </summary>
    public class SprTag
    {
        #region Properties

        public static DataRow DefaultRow
        {
            get
            {
                var tagTable = SprApplication.ActiveApplication.MdbDatabase.Tables["tag_data"];
                var returnRow = tagTable.NewRow();

                returnRow["tag_unique_id"] = 0;
                returnRow["tag_size"] = 0;
                returnRow["linkage_id_0"] = 0;
                returnRow["linkage_id_1"] = 0;
                returnRow["linkage_id_2"] = 0;
                returnRow["linkage_id_3"] = 0;
                returnRow["tag_text"] = string.Empty;
                returnRow["number_color"] = 0;
                returnRow["backgnd_color"] = 0;
                returnRow["leader_color"] = 0;
                returnRow["discipline"] = string.Empty;
                returnRow["creator"] = string.Empty;
                returnRow["computer_name"] = string.Empty;
                returnRow["status"] = string.Empty;

                return returnRow;
            }
        }

        internal DataRow Row
        {
            get
            {
                if (row == null)
                    row = GetTagRow();
                return row;
            }
        }
        private DataRow row;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     Holds the tag bitmask values used for tag placement.
        /// </summary>
        internal int Flags;

        /// <summary>
        ///     Indicates whether a leader will be present.
        /// </summary>
        public bool DisplayLeader
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & SprConstants.SprTagLeader) != 0;
            }
            set
            {
                // Set flag true/false
                if (value)
                    Flags |= SprConstants.SprTagLeader;
                else
                    Flags &= ~SprConstants.SprTagLeader;
            }
        }

        /// <summary>
        ///     Point where the tag is displayed.  If the tag has not been placed, the point coordinates will be 0, 0, 0.
        /// </summary>
        public SprPoint3D OriginPoint
        {
            get
            {
                if (!IsPlaced) return new SprPoint3D(0, 0, 0);
                return new SprPoint3D(Convert.ToDouble(Row["tag_origin_x"]),
                                      Convert.ToDouble(Row["tag_origin_y"]),
                                      Convert.ToDouble(Row["tag_origin_z"]));
            }
            set
            {
                if (!IsPlaced) return;
                Row["tag_origin_x"] = value.East;
                Row["tag_origin_y"] = value.North;
                Row["tag_origin_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Point for the end of the leader.  If the tag has not been placed, the point coordinates will be 0, 0, 0.
        /// </summary>
        public SprPoint3D LeaderPoint
        {
            get
            {
                if (!IsPlaced) return new SprPoint3D(0, 0, 0);
                return new SprPoint3D(Convert.ToDouble(Row["tag_point_x"]),
                                      Convert.ToDouble(Row["tag_point_y"]),
                                      Convert.ToDouble(Row["tag_point_z"]));
            }
            set
            {
                if (!IsPlaced) return;
                Row["tag_point_x"] = value.East;
                Row["tag_point_y"] = value.North;
                Row["tag_point_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Tag text.
        /// </summary>
        public string Text
        {
            get { return (string)Row["tag_text"]; }
            set { Row["tag_text"] = value; }
        }

        /// <summary>
        ///     Size of the tag bubble. If the tag has not been placed, the value is set to zero.
        /// </summary>
        public double Size
        {
            get { return IsPlaced ? Convert.ToDouble(Row["tag_size"]) : 0; }
            set { if (IsPlaced) Row["tag_size"] = value; }
        }

        /// <summary>
        ///     Date the tag was placed.  If the tag has not been placed, the value is N/A.
        /// </summary>
        public string DatePlaced
        {
            get { return IsPlaced ? Row["date_placed"].ToString() : "N/A"; }
            internal set { if (IsPlaced) Row["date_placed"] = value; }
        }

        /// <summary>
        ///     Date the tag was last edited.  If the tag has not been placed, the value is N/A.
        /// </summary>
        public string LastEdited
        {
            get { return IsPlaced ? Row["last_edited"].ToString() : "N/A"; }
            internal set { if (IsPlaced) Row["last_edited"] = value; }
        }

        /// <summary>
        ///     Color of the tag text.
        /// </summary>
        public Color TextColor
        {
            get { return SprUtilities.From0Bgr((int)Row["number_color"]); }
            set { Row["number_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Color of the tag background.
        /// </summary>
        public Color BackgroundColor
        {
            get { return SprUtilities.From0Bgr((int)Row["backgnd_color"]); }
            set { Row["backgnd_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Color of the tag leader line.
        /// </summary>
        public Color LeaderColor
        {
            get { return SprUtilities.From0Bgr((int)Row["leader_color"]); }
            set { Row["leader_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Discipline the tag is set to.
        /// </summary>
        public string Discipline
        {
            get { return (string)Row["discipline"]; }
            set { Row["discipline"] = value; }
        }

        /// <summary>
        ///     Author of the tag.
        /// </summary>
        public string Creator
        {
            get { return (string)Row["creator"]; }
            set { Row["creator"] = value; }
        }

        /// <summary>
        ///     Computer the tag was created on.
        /// </summary>
        public string ComputerName
        {
            get { return (string)Row["computer_name"]; }
            set { Row["computer_name"] = value; }
        }

        /// <summary>
        ///     Status of the tag.
        /// </summary>
        public string Status
        {
            get { return (string)Row["status"]; }
            set { Row["status"] = value; }
        }

        /// <summary>
        ///     Determines if the tag has been placed in SmartPlant Review.
        /// </summary>
        public bool IsPlaced { get; private set; }

        /// <summary>
        ///     Determines if the tag has an image stored in the MDB.
        /// </summary>
        public bool HasImage
        {
            get
            {
                if (!Row.Table.Columns.Contains("tag_image"))
                    return false;
                return Row["tag_image"] != DBNull.Value;
            }
        }

        /// <summary>
        ///     The stored image associated with the tag.
        /// </summary>
        public Image Image
        {
            get 
            {
                if (!HasImage)
                    return null;

                return DbMethods.GetDbImage(Row["tag_image"]);
            }
        }

        /// <summary>
        ///     Determines if the tag has labels linked to the MDB2
        /// </summary>
        public bool IsDataLinked
        {
            get { return Linkage.ToString() != "0 0 0 0"; }
        }

        /// <summary>
        ///     The object label linkage owned by the tag.
        /// </summary>
        public SprLinkage Linkage
        {
            get
            {
                var linkString = string.Format("{0} {1} {2} {3}",
                                      Row["linkage_id_0"],
                                      Row["linkage_id_1"],
                                      Row["linkage_id_2"],
                                      Row["linkage_id_3"]);

                return new SprLinkage(linkString);
            }
        }

        /// <summary>
        ///     Tag unique identification number.
        /// </summary>
        public int Id
        {
            get { return Convert.ToInt32(Row["tag_unique_id"]); }
            private set { Row["tag_unique_id"] = value; }
        }

        /// <summary>
        ///     The object the tag is attached to, if any.
        /// </summary>
        public SprObject LinkedObject
        {
            get
            {
                if (!IsDataLinked)
                    return null;

                if (linkedObject == null)
                    linkedObject = GetLinkedObject();

                return linkedObject;
            }
        }
        private SprObject linkedObject;

        #endregion

        #region Constructors

        public SprTag()
        {
            IsPlaced = false;

            Application = SprApplication.ActiveApplication;

            row = SprTag.DefaultRow;

            Id = 0;
            DisplayLeader = true;
            BackgroundColor = SprUtilities.From0Bgr(12632319);
            LeaderColor = SprUtilities.From0Bgr(12632319);
            Creator = Environment.GetEnvironmentVariable("USERNAME");
            ComputerName = Environment.GetEnvironmentVariable("COMPUTERNAME");
            Status = "Open";
        }
        internal SprTag(DataRow tagRow)
        {
            IsPlaced = tagRow["date_placed"] != DBNull.Value;

            Application = SprApplication.ActiveApplication;

            row = tagRow;
        }

        #endregion

        #region Methods
        
        private DataRow GetTagRow()
        {
            var result = Application.MdbDatabase.Tables["tag_data"].Rows.Find((object)Id);
            if (result == null)
                return SprTag.DefaultRow;

            IsPlaced = true;
            return result;
        }
        private SprObject GetLinkedObject()
        {
            var searchString = string.Format("FIND linkages = {0}", Linkage.ToString());
            var objIds = Application.ObjectDataSearch(searchString);
            if (objIds.Count > 1)
            {
                // Use the tag origin point to filter down to a single object
                var volumeFormat = "{0} N, {1} E, {2} El";
                var volumeStart = string.Format(volumeFormat, OriginPoint.North, OriginPoint.East, OriginPoint.Elevation);
                var volumeEnd = string.Format(volumeFormat, OriginPoint.North, OriginPoint.East, OriginPoint.Elevation);

                searchString = string.Format("{0}\nKEEP ONLY Volume Overlap {1} to {2}",
                                            searchString, volumeStart, volumeEnd);

                objIds = Application.ObjectDataSearch(searchString);
            }

            if (objIds.Count == 0)
                return null;
            else if (objIds.Count > 1)
                throw new SprException("Multiple objects found for linkage " + Linkage.ToString());
            return Application.GetObjectData(objIds[0]);
        }

        /// <summary>
        ///     Deletes the current tag from inside the SprApplication.
        /// </summary>
        /// <param name="setNext">Determines if the SprApplication.NextTag should be set to the deleted Id.</param>
        public void Delete(bool setNext = false)
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            // Delete the desired tag
            Application.SprStatus = Application.DrApi.TagDelete(Id, 0);

            // Set the deleted tag as the next tag number
            if (setNext)
                Application.NextTag = Id;

            // Update the main view
            Application.SprStatus = Application.DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Prompts a user to select new leader points for an existing tag.
        ///     DisplayLeader will automatically be set to true.
        /// </summary>
        public void EditLeader()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            if (!IsPlaced)
                throw new SprException("Tag {0} is not placed.", Id);

            var tagOrigin = new SprPoint3D();

            // Get an object on screen and set the origin point to its location
            var objId = Application.GetObjectId("SELECT NEW TAG START POINT", ref tagOrigin);
            if (objId == 0)
            {
                Application.TextWindow_Update("Tag placement canceled.");
                return;
            }

            // Get the tag leader point on screen
            var tagLeader = Application.GetPoint("SELECT NEW LEADER LOCATION", tagOrigin);
            if (tagLeader == null)
            {
                Application.TextWindow_Update("Tag placement canceled.");
                return;
            }

            var curObject = Application.GetObjectData(objId);

            DisplayLeader = true;
            Flags |= SprConstants.SprTagLabel;
            Flags |= SprConstants.SprTagEdit;

            // Update the tag with the new leader points
            Application.SprStatus = Application.DrApi.TagSetDbl(Id, 0, Flags, tagLeader.DrPointDbl,
                                                tagOrigin.DrPointDbl, curObject.Linkage.DrKey, Text);

            Refresh();

            // Flip the tag 180 degrees.  Intergraph is AWESOME!
            var swap = LeaderPoint;
            LeaderPoint = OriginPoint;
            OriginPoint = swap;

            Update();

            SendToTextWindow();
            Application.SprStatus = Application.DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Exports the saved tag snapshot if one exists.
        /// </summary>
        /// <param name="imagePath">The full path where the image will be saved.</param>
        /// <param name="overWrite">True if the destination image can be overwritten; Otherwise false. </param>
        public void ExportSnapshot(string imagePath, bool overWrite = true)
        {
            if (!HasImage)
                throw new SprException("No snapshot exists for the current image");

            if (Path.HasExtension(imagePath))
                imagePath = imagePath.Substring(0, imagePath.LastIndexOf('.'));

            if (Image.RawFormat.Equals(ImageFormat.Jpeg))
                imagePath += ".jpg";
            else if (Image.RawFormat.Equals(ImageFormat.Png))
                imagePath += ".png";
            else
                imagePath += ".bmp";

            if (!overWrite && File.Exists(imagePath))
                return;

            Image.Save(imagePath, Image.RawFormat);
        }

        /// <summary>
        ///     Locates the specified tag in the SmartPlant Review application main window.
        /// </summary>
        /// <param name="displayTag">Determines if the tag will be displayed in the main view.</param>
        public void Goto(bool displayTag = true)
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            // Update the text window with the tag information
            Application.TextWindow_Update(Row["tag_text"].ToString(), string.Format("Tag {0}", Id));

            // Locate the desired tag on the main screen with the specified visibility
            Application.SprStatus = Application.DrApi.GotoTag(Id, 0, Convert.ToInt32(displayTag));
        }

        /// <summary>
        ///     Prompts a user to place the current tag.
        /// </summary>
        public void Place()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            if (IsPlaced)
                throw new SprException("Tag {0} is already placed", Id);

            var tagOrigin = new SprPoint3D();

            // Get an object on screen and set the origin point to its location
            var objId = Application.GetObjectId("SELECT TAG START POINT", ref tagOrigin);
            if (objId == 0)
            {
                Application.TextWindow_Update("Tag placement canceled.");
                return;
            }

            // Get the tag leader point using the origin for depth
            var tagLeader = Application.GetPoint("SELECT TAG LEADER LOCATION", tagOrigin);
            if (tagLeader == null)
            {
                Application.TextWindow_Update("Tag placement canceled.");
                return;
            }

            var currentObject = Application.GetObjectData(objId);

            Flags |= SprConstants.SprTagLabel;

            // Set the tag registry values
            SprUtilities.SetTagRegistry(this);

            // Place the tag
            Application.SprStatus = Application.DrApi.TagSetDbl(Id, 0, Flags, ref tagLeader.DrPointDbl,
                                            ref tagOrigin.DrPointDbl, currentObject.Linkage.DrKey, Text);

            Update();

            // Clear the tag registry
            SprUtilities.ClearTagRegistry();

            SendToTextWindow();
            Application.SprStatus = Application.DrApi.ViewUpdate(1);
        }

        /// <summary>
        ///     Loads the latest tag information from the MDB database.
        /// </summary>
        public void Refresh()
        {
            var updatedTable = DbMethods.GetDbTable(Application.MdbPath, Row.Table.TableName);
            var updatedRow = updatedTable.Rows.Find((object)Id);
            
            row.ItemArray = updatedRow.ItemArray;
            row.AcceptChanges();
        }

        /// <summary>
        ///     Updates the SmartPlant Review text window with the SprTag text.
        /// </summary>
        public void SendToTextWindow()
        {
            Application.TextWindow_Update(Text, string.Format("Tag {0}", Id));
        }

        /// <summary>
        ///     Converts the current tag to a string representation.
        /// </summary>
        public override string ToString()
        {
            return string.Format("Tag {0}: {1}", Id, Text);
        }

        /// <summary>
        ///     Converts the current tag to a string representation using the specified format.
        /// </summary>
        public string ToString(string format)
        {
            if (format == string.Empty)
                return ToString();
            else
                return string.Format(format, Row.ItemArray);
        }

        /// <summary>
        ///     Saves a snapshot of the current SprApplication main view to the MDB database.
        /// </summary>
        /// <param name="snapShot">The snapshot format the image will be created with.</param>
        /// <param name="ZoomToTag">Determines if the main view zooms to the tag in the main SmartPlant screen.</param>
        public void TakeSnapshot(bool ZoomToTag, SprSnapShot snapShot = null)
        {
            if (!DbMethods.AddDbField(Application.MdbPath, "tag_data", "tag_image", "OLEOBJECT"))
                throw new SprException("Error creating tag image field in MDB database");

            Refresh();

            if (ZoomToTag)
                Goto();

            var snap = snapShot ?? Application.DefaultSnapshot;   
            var image = Application.TakeSnapshot("dbImage_temp", SprSnapShot.TempDirectory, snap);
            var ms = new MemoryStream();
            image.Save(ms, image.RawFormat);

            Row["tag_image"] = ms.ToArray();

            Update();

            File.Delete(Path.Combine(SprSnapShot.TempDirectory, "dbImage_temp"));
        }

        /// <summary>
        ///     Updates the MDB database with the current tag information.
        /// </summary>
        public void Update()
        {
            var filter = string.Format("tag_unique_id = {0}", Id);
            DbMethods.UpdateDbTable(Application.MdbPath, filter, Row.Table);
        }

        #endregion
    }
}