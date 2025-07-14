namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public static class Constants
{
    // View names
    public const string FrontView = "Front_view";
    public const string SideView = "Side_view";
    public const string TopView = "Top_view";
    public const string DetailView = "Detail_view";
    public const string SectionView = "Section_view";
    //public const string OverlaySectionView = "Section View O-O";
    public const string OverlaySectionView = "Section View L-L";
    //public const string OverlayDetailView = "Drawing View6";
    public const string OverlayDetailView = "Drawing View4";
    public const string OverlaySideView = "Drawing View1";
    public const string OverlayTopView = "Drawing View3";
    public const string OverlayFrontView = "Drawing View2";
    public const string OverlaySideView2 = "Drawing View7";
    //public const string OverlayDetailView2 = "Detail View P (100 : 1)";
    /// </summary>
    // Sketch names
    public const string SketchEngraving = "sketch_engraving";
    public const string SketchGrooveDimensions = "sketch_groove_dimensions";

    // Layer names
    public const string FormatLayer = "FORMAT";

    // Property names
    public const string EngravedTextProperty = "Engraved Text";

    // Table identifiers
    public const string DimensionTable = "dimension";
    public const string HowToOrderTable = "how_to_order";
    public const string LabelAsTable = "label_as";
    public const string PolishTable = "polish";
    public const string CoiningNote = "coining_note";
    // Misc
    public const string DatumFeatureLabel = "A";

    public static class ConfigKeys
    {
        public const string ScalingFSV = "scaling_fsv";
        public const string ScalingDSV = "scaling_dsv";
        public const string FrontViewPosX = "front_view_posX";
        public const string FrontViewPosY = "front_view_posY";
        public const string SideViewDX = "side_view_dX";
        public const string SideViewDY = "side_view_dY";
        public const string TopViewDX = "top_view_dX";
        public const string TopViewDY = "top_view_dY";
        public const string DetailViewPosX = "detail_view_posX";
        public const string DetailViewPosY = "detail_view_posY";
        public const string SectionViewPosX = "section_view_posX";
        public const string Material = "material";
        public const string Author = "author";
        public const string Packaging = "packaging";
        public const string Engrave = "engrave";
        public const string PolishText = "polish_text";
        public const string DimensionKeysInTable = "dimension_keys_in_table";
    }
}