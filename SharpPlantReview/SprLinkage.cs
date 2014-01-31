//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Holds the linkage values for retrieving object and label information.
    /// </summary>
    public class SprLinkage
    {
        #region SprLinkage Properties

        /// <summary>
        ///     The active COM reference to the DrKey class.
        /// </summary>
        internal dynamic DrKey;

        /// <summary>
        ///     Linkage_Id_0
        /// </summary>
        public int Id1
        { 
            get { return DrKey.LabelKey1; }
            private set { DrKey.LabelKey1 = value; }
        }

        /// <summary>
        ///     Linkage_Id_1
        /// </summary>
        public int Id2
        {
            get { return DrKey.LabelKey2; }
            private set { DrKey.LabelKey2 = value; }
        }

        /// <summary>
        ///     Linkage_Id_2
        /// </summary>
        public int Id3
        {
            get { return DrKey.LabelKey3; }
            private set { DrKey.LabelKey3 = value; }
        }

        /// <summary>
        ///     Linkage_Id_3
        /// </summary>
        public int Id4
        {
            get { return DrKey.LabelKey4; }
            private set { DrKey.LabelKey4 = value; }
        }

        #endregion

        // Constructors
        public SprLinkage()
        {
            DrKey = Activator.CreateInstance(SprImportedTypes.DrKey);
        }
        public SprLinkage(string linkage) : this()
        {
            var links = linkage.Split(' ');
            if (links.Length != 4)
                throw new SprException("This link is bullshit.");

            Id1 = Int32.Parse(links[0]);
            Id2 = Int32.Parse(links[1]);
            Id3 = Int32.Parse(links[2]);
            Id4 = Int32.Parse(links[3]);
        }
        public SprLinkage(int link1, int link2, int link3, int link4) : this()
        {
            Id1 = link1;
            Id2 = link2;
            Id3 = link3;
            Id4 = link4;
        }
        public SprLinkage(dynamic drKey)
        {
            DrKey = drKey;
        }

        /// <summary>
        ///     Returns a string representing the linkage values.
        /// </summary>
        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3}", Id1, Id2, Id3, Id4);
        }
    }
}
