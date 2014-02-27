//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the properties for controlling text annotations in SmartPlant Review.
    /// </summary>
    public class SprAnnotation : SprDbObject
    {
        #region Properties

        /// <summary>
        ///     The active COM reference to the DrAnnotationDbl class
        /// </summary>
        internal dynamic DrAnnotationDbl
        {
            get
            {
                if (Data == null) return null;
                dynamic drAnno = Activator.CreateInstance(SprImportedTypes.DrAnnotationDbl);
                drAnno.BackgroundColor = Data["bg_color"];
                drAnno.LineColor = Data["line_color"];
                drAnno.TextColor = Data["text_color"];
                drAnno.CenterPoint = CenterPoint.DrPointDbl;
                drAnno.LeaderPoint = LeaderPoint.DrPointDbl;
                drAnno.Flags = Flags;
                drAnno.Text = Text;

                return drAnno;
            }
            set
            {
                if (Data == null) return;
                Data["bg_color"] = value.BackgroundColor;
                Data["line_color"] = value.LineColor;
                Data["text_color"] = value.TextColor;
                CenterPoint = new SprPoint3D(value.CenterPoint.East, value.CenterPoint.North, value.CenterPoint.Elevation);
                LeaderPoint = new SprPoint3D(value.LeaderPoint.East, value.LeaderPoint.North, value.LeaderPoint.Elevation);
                Flags = value.Flags;
                Text = value.Text;
            }
        }

        /// <summary>
        ///     Annotation background color if visible (0BGR format).
        /// </summary>
        public Color BackgroundColor
        {
            get { return SprUtilities.From0Bgr((int)Data["bg_color"]); }
            set
            { 
                Data["bg_color"] = SprUtilities.Get0Bgr(value);
                if (DrAnnotationDbl != null) DrAnnotationDbl.BackgroundColor = Data["bg_color"];
            }
        }

        /// <summary>
        ///     Leader line color.
        /// </summary>
        public Color LineColor
        {
            get { return SprUtilities.From0Bgr((int)Data["line_color"]); }
            set
            {
                Data["line_color"] = SprUtilities.Get0Bgr(value);
                if (DrAnnotationDbl != null) DrAnnotationDbl.LineColor = Data["line_color"];
            }
        }

        /// <summary>
        ///     Annotation text color.
        /// </summary>
        public Color TextColor
        {
            get { return SprUtilities.From0Bgr((int)Data["text_color"]); }
            set
            {
                Data["text_color"] = SprUtilities.Get0Bgr(value);
                if (DrAnnotationDbl != null) DrAnnotationDbl.TextColor = Data["text_color"];
            }
        }

        /// <summary>
        ///     Center point of the annotation object.
        /// </summary>
        public SprPoint3D CenterPoint
        {
            get
            {
                if (!IsPlaced) return new SprPoint3D(0, 0, 0);
                //return 

                return new SprPoint3D(Convert.ToDouble(Data["center_x"]),
                                      Convert.ToDouble(Data["center_y"]),
                                      Convert.ToDouble(Data["center_z"]));
            }
            set
            {
                if (!IsPlaced) return;
                Data["center_x"] = value.East;
                Data["center_y"] = value.North;
                Data["center_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     End point of the leader line.
        /// </summary>
        public SprPoint3D LeaderPoint
        {
            get
            {
                if (!IsPlaced) return new SprPoint3D(0, 0, 0);
                return new SprPoint3D(Convert.ToDouble(Data["leader_x"]),
                                      Convert.ToDouble(Data["leader_y"]),
                                      Convert.ToDouble(Data["leader_z"]));
            }
            set
            {
                var pt = new SprPoint3D(value);
                Data["leader_x"] = value.East;
                Data["leader_y"] = value.North;
                Data["leader_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Indicates whether a leader will be present.
        /// </summary>
        public bool DisplayLeader
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & SprConstants.SprAnnoLeader) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= SprConstants.SprAnnoLeader;
                else Flags &= ~SprConstants.SprAnnoLeader;
            }
        }

        /// <summary>
        ///     Holds the tag bitmask values used for annotation placement.
        /// </summary>
        internal int Flags
        {
            get { return DrAnnotationDbl.Flags; }
            set { DrAnnotationDbl.Flags = value; }
        }

        /// <summary>
        ///     Indicates whether an arrowhead will be present.
        /// </summary>
        public bool DisplayArrowhead
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & SprConstants.SprAnnoArrow) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= SprConstants.SprAnnoArrow;
                else Flags &= ~SprConstants.SprAnnoArrow;
            }
        }

        /// <summary>
        ///     Indicates whether the background will be displayed.
        /// </summary>
        public bool DisplayBackground
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & SprConstants.SprAnnoBackground) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= SprConstants.SprAnnoBackground;
                else Flags &= ~SprConstants.SprAnnoBackground;
            }
        }

        /// <summary>
        ///     Indicates that the annotation should be maintained between sessions.
        /// </summary>
        public bool Persistent
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & SprConstants.SprAnnoPersist) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= SprConstants.SprAnnoPersist;
                else Flags &= ~SprConstants.SprAnnoPersist;
            }
        }

        /// <summary>
        ///     The annotation text string.
        /// </summary>
        public string Text
        {
            get { return Row["text_string"].ToString(); }
            set { Row["text_string"] = value; }
        }
            
        /// <summary>
        ///     The object associated with the annotation.
        /// </summary>
        //public SprObject AssociatedObject
        //{
        //    get
        //    {
        //        try { return Application.GetObjectData(Convert.ToInt32(Data["object_id"])); }
        //        catch (KeyNotFoundException) { return null; }
        //    }
        //    internal set
        //    {
        //        try { Data["object_id"] = value.Id; }
        //        catch (KeyNotFoundException) { }
        //    }
        //}

        /// <summary>
        ///     Annotation type.
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        ///     The full information profile of the current annotation.  Controls all the annotation properties.
        /// </summary>
        public Dictionary<string, object> Data { get; set; }

        public bool IsPlaced { get; internal set; }

        #endregion

        #region Constructors

        public SprAnnotation() : base()
        {
            // Create the backing Annotation object
            DrAnnotationDbl = Activator.CreateInstance(SprImportedTypes.DrAnnotationDbl);

            IsPlaced = false;
            DisplayLeader = true;
            DisplayBackground = true;
            Type = "Standard";
        }
        public SprAnnotation(DataRow dataRow) : base(dataRow)
        {
            // Create the backing Annotation object
            DrAnnotationDbl = Activator.CreateInstance(SprImportedTypes.DrAnnotationDbl);

            IsPlaced = dataRow["date_placed"] != DBNull.Value;
            DisplayLeader = (Convert.ToDouble(Row["leader_x"]) != 0); // Or something
        }

        #endregion

        #region Methods

        protected override DataRow GetDefaultRow()
        {
            var annoTable = SprApplication.ActiveApplication.MdbDatabase.Tables["text_annotations"];
            var annoRow = annoTable.NewRow();
            annoRow["id"] = 0;
            annoRow["type_id"] = 0;
            annoRow["bg_color"] = 12632319;
            annoRow["line_color"] = 12632319;
            annoRow["text_color"] = 0;
            annoRow["text_string"] = string.Empty;

            return annoRow;
        }
        private Image GetImage()
        {
            return null;
        }
        private SprPoint3D GetCenterPoint()
        {
            return null;
        }
        private SprPoint3D GetLeaderPoint()
        {
            return null;
        }


        /// <summary>
        ///     Prompts a user to select new leader points for an existing annotation.
        ///     DisplayLeader will automatically be set to true.
        /// </summary>
        ///TODO
        public void EditLeader()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            if (!IsPlaced)
                throw new SprException("Annotation {0} is not placed.", Id);

            var annoOrigin = new SprPoint3D();

            // Get an object on screen and set the origin point to its location
            var objId = Application.GetObjectId("SELECT NEW ANNOTATION START POINT", ref annoOrigin);
            if (objId == 0)
            {
                Application.Windows.TextWindow.Text = "Annotation placement canceled.";
                return;
            }

            // Get the annotation leader point on screen
            var annoLeader = Application.GetPoint("SELECT NEW LEADER LOCATION", annoOrigin);
            if (annoLeader == null)
            {
                Application.Windows.TextWindow.Text = "Tag placement canceled.";
                return;
            }

            var curObject = Application.GetObjectData(objId);

            DisplayLeader = true;
            Flags |= SprConstants.SprTagLabel;
            Flags |= SprConstants.SprTagEdit;

            // Update the tag with the new leader points
            Application.SprStatus = Application.DrApi.TagSetDbl(Id, 0, Flags, annoLeader.DrPointDbl,
                                                annoOrigin.DrPointDbl, curObject.Linkage.DrKey, Text);

            //Refresh();

            //// Flip the tag 180 degrees.  Intergraph is AWESOME!
            //var swap = LeaderPoint;
            //LeaderPoint = OriginPoint;
            //OriginPoint = swap;

            Update();

            SendToTextWindow();
            //Application.SprStatus = Application.DrApi.ViewUpdate(1);
            //Application.Run(SprNativeMethods.ViewUpdate, 1);
        }


        /// <summary>
        ///     Prompts a user to place the current annotation.
        /// </summary>
        /// TODO
        public void Place()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            if (IsPlaced)
                throw new SprException("Annotation {0} is already placed", Id);

            var leaderPoint = new SprPoint3D();

            // Get the annotation leader point
            int objId = Application.GetObjectId("SELECT A POINT ON AN OBJECT TO LOCATE THE ANNOTATION", ref leaderPoint);
            if (objId == 0)
            {
                Application.Windows.TextWindow.Text = "Annotation placement canceled.";
                return;
            }

            // Get the annotation center point using the leaderpoint for depth calculation
            var centerPoint = Application.GetPoint("SELECT THE CENTER POINT FOR THE ANNOTATION LABEL", leaderPoint);
            if (centerPoint == null)
            {
                Application.Windows.TextWindow.Text = "Annotation placement canceled.";
                return;
            }

            // Set the annotation points
            LeaderPoint = leaderPoint;
            CenterPoint = centerPoint;

            // Place the annotation on screen
            var annoId = 0;
            //Application.SprStatus = Application.DrApi.AnnotationCreateDbl(Type, ref DrAnnotationDbl, out annoId);

            // Link the located object to the annotation
            //SprStatus = DrApi.AnnotationDataSet(annoId, anno.Type, ref drAnno, ref objId);

            Refresh();
            // Retrieve the placed annotation data
            //anno = Annotations_Get(anno.Id);

            // Add an ObjectId field
            //Annotations_AddDataField("object_id");

            // Save the ObjectId to the annotation data
            //anno.Data["object_id"] = objId;

            // Update the annotation
            //Annotations_Update(anno);

            // Update the text window
            Application.Windows.TextWindow.Title = string.Format("Annotation {0}", Id);
            Application.Windows.TextWindow.Text = Text;

            // Update the main view
            Application.SprStatus = Application.DrApi.ViewUpdate(1);

        }

        /// <summary>
        ///     Updates the SmartPlant Review text window with the SprAnnotation text.
        /// </summary>
        public void SendToTextWindow()
        {
            Application.Windows.TextWindow.Title = string.Format("Tag {0}", Id);
            Application.Windows.TextWindow.Text = Text;
        }

        /// <summary>
        ///     Converts the current annotation to a string representation.
        /// </summary>
        public override string ToString()
        {
            return string.Format("Annotation {0}: {1}", Id, Text);
        }

        #endregion
    }
}