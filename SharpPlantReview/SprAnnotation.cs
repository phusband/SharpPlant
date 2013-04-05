//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Drawing;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the properties for controlling text annotations in SmartPlant Review.
    /// </summary>
    public class SprAnnotation
    {
        #region Annotation Properties

        /// <summary>
        ///     Active COM reference to the DrAnnotationDbl class.
        /// </summary>
        internal dynamic DrAnnotationDbl;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     Annotation background color if visible (0BGR format).
        /// </summary>
        public Color BackgroundColor
        {
            get { return IsActive ? SprUtilities.From0Bgr(DrAnnotationDbl.BackgroundColor) : Color.Empty; }
            set { if (IsActive) DrAnnotationDbl.BackgroundColor = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Leader line color.
        /// </summary>
        public Color LineColor
        {
            get { return IsActive ? SprUtilities.From0Bgr(DrAnnotationDbl.LineColor) : Color.Empty; }
            set { if (IsActive) DrAnnotationDbl.LineColor = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Annotation text color.
        /// </summary>
        public Color TextColor
        {
            get { return IsActive ? SprUtilities.From0Bgr(DrAnnotationDbl.TextColor) : Color.Empty; }
            set { if (IsActive) DrAnnotationDbl.TextColor = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Center point of the annotation object.
        /// </summary>
        public SprPoint3D CenterPoint
        {
            get
            {
                return IsActive ? new SprPoint3D(DrAnnotationDbl.CenterPoint) : null;
            }
            set
            {
                if (!IsActive) return;
                DrAnnotationDbl.CenterPoint.East = value.East;
                DrAnnotationDbl.CenterPoint.North = value.North;
                DrAnnotationDbl.CenterPoint.Elevation = value.Elevation;
            }
        }

        /// <summary>
        ///     End point of the leader line.
        /// </summary>
        public SprPoint3D LeaderPoint
        {
            get
            {
                return IsActive ? new SprPoint3D(DrAnnotationDbl.LeaderPoint) : null;
            }
            set
            {
                if (!IsActive) return;
                DrAnnotationDbl.LeaderPoint.East = value.East;
                DrAnnotationDbl.LeaderPoint.North = value.North;
                DrAnnotationDbl.LeaderPoint.Elevation = value.Elevation;
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
                return (DrAnnotationDbl.Flags & SprConstants.SprAnnoLeader) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= SprConstants.SprAnnoLeader;
                else DrAnnotationDbl.Flags &= ~SprConstants.SprAnnoLeader;
            }
        }

        /// <summary>
        ///     Indicates whether an arrowhead will be present.
        /// </summary>
        public bool DisplayArrowhead
        {
            get
            {
                // Return the bitwise zero check
                return (DrAnnotationDbl.Flags & SprConstants.SprAnnoArrow) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= SprConstants.SprAnnoArrow;
                else DrAnnotationDbl.Flags &= ~SprConstants.SprAnnoArrow;
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
                return (DrAnnotationDbl.Flags & SprConstants.SprAnnoBackground) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= SprConstants.SprAnnoBackground;
                else DrAnnotationDbl.Flags &= ~SprConstants.SprAnnoBackground;
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
                return (DrAnnotationDbl.Flags & SprConstants.SprAnnoPersist) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= SprConstants.SprAnnoPersist;
                else DrAnnotationDbl.Flags &= ~SprConstants.SprAnnoPersist;
            }
        }

        /// <summary>
        ///     Annotation text.
        /// </summary>
        public string Text
        {
            get { return IsActive ? DrAnnotationDbl.Text : null; }
            set { if (IsActive) DrAnnotationDbl.Text = value; }
        }

        /// <summary>
        ///     The session unique ObjectID of the annotation.
        /// </summary>
        public int AnnotationId { get; internal set; }

        /// <summary>
        ///     The object associated with the annotation.
        /// </summary>
        public SprObjectData AssociatedObject { get; internal set; }

        /// <summary>
        ///     Annotation type.
        /// </summary>
        public string Type { get; set; }

        // Determines if a reference to the COM object is established
        private bool IsActive
        {
            get { return DrAnnotationDbl != null; }
        }

        #endregion

        // Annotation constructor
        public SprAnnotation()
        {
            // Link the parent application
            Application = SprApplication.ActiveApplication;

            // Get a new DrAnnotationDbl object
            DrAnnotationDbl = Activator.CreateInstance(SprImportedTypes.DrAnnotationDbl);

            // Set the default flags
            DrAnnotationDbl.Flags |= SprConstants.SprAnnoLeader | SprConstants.SprAnnoBackground;

            // Set the default annotation type
            Type = "Standard";

            // Set the default colors
            BackgroundColor = SprUtilities.From0Bgr(8454143);
            LineColor = SprUtilities.From0Bgr(8454143);
            TextColor = SprUtilities.From0Bgr(0);
        }
    }
}