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
    public class SprTag : SprDbObject, IDisposable
    {
        #region Properties

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
            get { return _originPoint ?? (_originPoint = GetOriginPoint()); }
            set { SetOriginPoint(value); }
        }
        private SprPoint3D _originPoint;

        /// <summary>
        ///     Point for the end of the leader.  If the tag has not been placed, the point coordinates will be 0, 0, 0.
        /// </summary>
        public SprPoint3D LeaderPoint
        {
            get { return _leaderPoint ?? (_leaderPoint = GetLeaderPoint()); }
            set { SetLeaderPoint(value); }
        }
        private SprPoint3D _leaderPoint;

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
        public String DatePlaced
        {
            get
            {
                return Row["date_placed"] == DBNull.Value ? Row["date_placed"].ToString() : "N/A";
                //return date = string.erm (date = DateTime.;
            }
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
        public bool IsPlaced //{ get; private set; }
        {
            get { return (Convert.ToInt32(Row["tag_size"])) != 0; }
        }

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
            get { return _image ?? (_image = GetImage()); }
        }
        private Image _image;

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
        ///     The object the tag is attached to, if any.
        /// </summary>
        public SprObject LinkedObject
        {
            get { return _linkedObject ?? (_linkedObject = GetLinkedObject()); }
        }
        private SprObject _linkedObject;

        #endregion

        public SprTagCollection Collection
        {
            get { return _collection; }
            internal set { _collection = value; }
        }
        private SprTagCollection _collection;

        #region Constructors

        public SprTag()
        {
            DisplayLeader = true;
        }
        public SprTag(DataRow dataRow) : base(dataRow)
        {
            DisplayLeader = (Convert.ToDouble(Row["tag_point_x"]) != 0); // Or something
        }

        #endregion

        #region IDisposable

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (HasImage)
                        Image.Dispose();
                }

                base.Dispose(disposing);
            }
        }

        #endregion

        #region Methods

        protected override DataRow GetDataRow()
        {
            if (Collection == null)
                throw new SprException("This tag is not part of a collection");

            var dataRow = Collection.Table.Rows.Find(Id);
            return dataRow ?? (DefaultRow);
        }
        protected override DataRow GetDefaultRow()
        {
            //var tagTable = Collection.Table;
            //var tagTable = SprApplication.ActiveApplication.Tags.Table;
            var tagTable = Application.MdbDatabase.Tables[SprConstants.MdbTagTable];
            var tagRow = tagTable.NewRow();
            tagRow["tag_unique_id"] = 0;
            tagRow["tag_size"] = 0; // Controls IsPlaced
            tagRow["linkage_id_0"] = 0;
            tagRow["linkage_id_1"] = 0;
            tagRow["linkage_id_2"] = 0;
            tagRow["linkage_id_3"] = 0;
            tagRow["tag_text"] = string.Empty;
            tagRow["number_color"] = 0;
            tagRow["backgnd_color"] = 12632319;
            tagRow["leader_color"] = 12632319;
            tagRow["discipline"] = string.Empty;
            tagRow["creator"] = Environment.GetEnvironmentVariable("USERNAME");
            tagRow["computer_name"] = Environment.GetEnvironmentVariable("COMPUTERNAME");
            tagRow["status"] = "Open";

            return tagRow;
        }

        private Image GetImage()
        {
            return !HasImage ? null : DbMethods.GetDbImage(Row["tag_image"]);
        }
        private SprObject GetLinkedObject()
        {
            if (!IsDataLinked)
                return null;

            var searchString = string.Format("FIND linkages = {0}", Linkage);
            var objIds = Application.ObjectDataSearch(searchString);
            if (objIds.Count > 1)
            {
                // Use the tag origin point to filter down to a single object
                const string volumeFormat = "{0} N, {1} E, {2} El";
                var volumeStart = string.Format(volumeFormat, OriginPoint.North, OriginPoint.East, OriginPoint.Elevation);
                var volumeEnd = string.Format(volumeFormat, OriginPoint.North, OriginPoint.East, OriginPoint.Elevation);

                searchString = string.Format("{0}\nKEEP ONLY Volume Overlap {1} to {2}",
                                            searchString, volumeStart, volumeEnd);

                objIds = Application.ObjectDataSearch(searchString);
            }

            if (objIds.Count == 0)
                return null;
            if (objIds.Count > 1)
                throw new SprException("Multiple objects found for linkage {0}", Linkage);

            return Application.GetObjectData(objIds[0]);
        }
        private SprPoint3D GetLeaderPoint()
        {
            if (!IsPlaced)
                return null;

            return new SprPoint3D(Convert.ToDouble(Row["tag_point_x"]),
                                  Convert.ToDouble(Row["tag_point_y"]),
                                  Convert.ToDouble(Row["tag_point_z"]));
        }
        private SprPoint3D GetOriginPoint()
        {
            if (!IsPlaced)
                return null;

            return new SprPoint3D(Convert.ToDouble(Row["tag_origin_x"]),
                                  Convert.ToDouble(Row["tag_origin_y"]),
                                  Convert.ToDouble(Row["tag_origin_z"]));
        }

        private void SetLeaderPoint(SprPoint3D newpoint)
        {
            if (!IsPlaced)
                throw SprExceptions.SprTagNotPlaced;

            _leaderPoint = newpoint;
            Row["tag_origin_x"] = newpoint.East;
            Row["tag_origin_y"] = newpoint.North;
            Row["tag_origin_z"] = newpoint.Elevation;
        }
        private void SetOriginPoint(SprPoint3D newpoint)
        {
            if (!IsPlaced)
                throw SprExceptions.SprTagNotPlaced;

            _originPoint = newpoint;
            Row["tag_origin_x"] = newpoint.East;
            Row["tag_origin_y"] = newpoint.North;
            Row["tag_origin_z"] = newpoint.Elevation;
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
            if (Application.SprStatus != 0)
                throw Application.SprException;

            // Set the deleted tag as the next tag number
            if (setNext)
                Application.NextTag = Id;

            // Update the main view
            Application.SprStatus = Application.DrApi.ViewUpdate(1);
            if (Application.SprStatus != 0)
                throw Application.SprException;

            Collection.Remove(this);
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
            Application.HighlightClear();
            var objId = Application.GetObjectId("SELECT NEW TAG START POINT", ref tagOrigin);
            if (objId == 0)
            {
                Application.Windows.TextWindow.Text = "Tag placement canceled.";
                return;
            }

            // Highlight the selected object
            Application.HighlightObject(objId, Color.Fuchsia);

            // Get the tag leader point on screen
            var tagLeader = Application.GetPoint("SELECT NEW LEADER LOCATION", tagOrigin);
            if (tagLeader == null)
            {
                Application.HighlightClear();
                Application.Windows.TextWindow.Text = "Tag placement canceled.";
                return;
            }

            _linkedObject = Application.GetObjectData(objId);

            DisplayLeader = true;
            Flags |= SprConstants.SprTagLabel;
            Flags |= SprConstants.SprTagEdit;

            // Update the tag with the new leader points
            Application.SprStatus = Application.DrApi.TagSetDbl(Id, 0, Flags, tagLeader.DrPointDbl,
                                                tagOrigin.DrPointDbl, LinkedObject.Linkage.DrKey, Text);
            if (Application.SprStatus != 0)
                throw Application.SprException;

            Refresh();

            // Flip the tag 180 degrees.  Intergraph is AWESOME!
            var swap = LeaderPoint;
            LeaderPoint = OriginPoint;
            OriginPoint = swap;

            Update();

            SendToTextWindow();
            Application.HighlightClear();

            Application.SprStatus = Application.DrApi.ViewUpdate(1);
            if (Application.SprStatus != 0)
                throw Application.SprException;
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

            // Strip off the extension
            if (Path.HasExtension(imagePath))
                imagePath = imagePath.Substring(0, imagePath.LastIndexOf('.'));

            // Get the image encoding
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
            Application.Windows.TextWindow.Title = string.Format("Tag {0}", Id);
            Application.Windows.TextWindow.Text = Row["tag_text"].ToString();

            // Locate the desired tag on the main screen with the specified visibility
            Application.SprStatus = Application.DrApi.GotoTag(Id, 0, Convert.ToInt32(displayTag));
            if (Application.SprStatus != 0)
                throw Application.SprException;
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
            Application.HighlightClear();
            var objId = Application.GetObjectId("SELECT TAG START POINT", ref tagOrigin);
            if (objId == 0)
            {
                Application.Windows.TextWindow.Text = "Tag placement canceled.";
                return;
            }

            // Highlight the selected object
            Application.HighlightObject(objId, Color.Fuchsia);

            // Get the tag leader point using the origin for depth
            var tagLeader = Application.GetPoint("SELECT TAG LEADER LOCATION", tagOrigin);
            if (tagLeader == null)
            {
                Application.HighlightClear();
                Application.Windows.TextWindow.Text = "Tag placement canceled.";
                return;
            }

            _linkedObject = Application.GetObjectData(objId);

            Flags |= SprConstants.SprTagLabel;
            Id = Application.NextTag;

            // Set the tag registry values
            SprUtilities.SetTagRegistry(this);

            // Place the tag
            Application.SprStatus = Application.DrApi.TagSetDbl(Id, 0, Flags, ref tagLeader.DrPointDbl,
                                            ref tagOrigin.DrPointDbl, LinkedObject.Linkage.DrKey, Text);
            if (Application.SprStatus != 0)
                throw Application.SprException;

            //IsPlaced = true;
            _leaderPoint = tagLeader;
            _originPoint = tagOrigin;

            Application.Tags.Add(this);
            Refresh();

            // Clear the tag registry
            SprUtilities.ClearTagRegistry();

            SendToTextWindow();

            Application.HighlightClear();
            Application.SprStatus = Application.DrApi.ViewUpdate(1);
            if (Application.SprStatus != 0)
                throw Application.SprException;
        }

        /// <summary>
        ///     Updates the SmartPlant Review text window with the SprTag text.
        /// </summary>
        public void SendToTextWindow()
        {
            Application.Windows.TextWindow.Title = string.Format("Tag {0}", Id);
            Application.Windows.TextWindow.Text = Text;
        }

        /// <summary>
        ///     Converts the current tag to a string representation.
        /// </summary>
        public override string ToString()
        {
            return string.Format("Tag {0}: {1}", Id, Text);
        }

        /// <summary>
        ///     Saves a snapshot of the current SprApplication main view to the MDB database.
        /// </summary>
        /// <param name="snapShot">The snapshot format the image will be created with.</param>
        /// <param name="zoomToTag">Determines if the main view zooms to the tag in the main SmartPlant screen.</param>
        public void TakeSnapshot(bool zoomToTag, SprSnapShot snapShot = null)
        {
            if (!Row.Table.Columns.Contains("tag_image"))
            {
                Application.Tags.AddDataField("tag_image", "OLEOBJECT");
                var oldVals = Row.ItemArray;
                Refresh();
                Row.ItemArray = oldVals;
            }

            if (zoomToTag)
                Goto();

            var snap = snapShot ?? Application.DefaultSnapshot;
            var tempName = string.Format("dbImage_tag_{0}", Id);

            var ms = new MemoryStream();
            _image = Application.TakeSnapshot(tempName, SprSnapShot.TempDirectory, snap);
            _image.Save(ms, _image.RawFormat);

            Row["tag_image"] = ms.ToArray();
            ms.Dispose();

            Update();
        }

        /// <summary>
        ///     Deletes the snapshot owned by the current <see cref="SprTag"/>.
        /// </summary>
        public void DeleteSnapshot()
        {
            if (Image != null)
                Image.Dispose();

            Row["tag_image"] = DBNull.Value;

            Update();
        }

        #endregion
    }
}