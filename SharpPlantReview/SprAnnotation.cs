//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Drawing;
using System.Collections.Generic;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the properties for controlling text annotations in SmartPlant Review.
    /// </summary>
    public class SprAnnotation
    {
        #region Annotation Properties

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
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     Annotation background color if visible (0BGR format).
        /// </summary>
        public Color BackgroundColor
        {
            get { return SprUtilities.From0Bgr((int) Data["bg_color"]); }
            set { Data["bg_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Leader line color.
        /// </summary>
        public Color LineColor
        {
            get { return SprUtilities.From0Bgr((int) Data["line_color"]); }
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
        internal int Flags { get; set; }

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
            get { return Data["text_string"].ToString(); }
            set { Data["text_string"] = value; }
        }

        /// <summary>
        ///     The session unique ID of the annotation.
        /// </summary>
        public int Id
        {
            get { return Convert.ToInt32(Data["id"]); }
            internal set { Data["id"] = value; }
        }
            
        /// <summary>
        ///     The object associated with the annotation.
        /// </summary>
        public SprObjectData AssociatedObject
        {
            get
            {
                try { return Application.GetObjectData(Convert.ToInt32(Data["object_id"])); }
                catch (KeyNotFoundException) { return null; }
            }
            internal set
            {
                try { Data["object_id"] = value.ObjectId; }
                catch (KeyNotFoundException) { }
            }
        }

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

        // Annotation constructor
        public SprAnnotation()
        {
            // Link the parent application
            Application = SprApplication.ActiveApplication;

            // Create the backing Annotation object
            DrAnnotationDbl = Activator.CreateInstance(SprImportedTypes.DrAnnotationDbl);

            // Set as not placed by default
            IsPlaced = false;

            // Set the default flags
            Flags |= SprConstants.SprAnnoLeader | SprConstants.SprAnnoBackground;

            // Create a new data dictionary from the template
            Data = SprUtilities.AnnotationTemplate();

            // Set the tag to the next available tag number
            Id = Application.NextAnnotation;

            // Set the default annotation type
            Type = "Standard";

            // Set the default colors
            BackgroundColor = SprUtilities.From0Bgr(8454143);
            LineColor = SprUtilities.From0Bgr(8454143);
            TextColor = SprUtilities.From0Bgr(0);
        }
    }
}