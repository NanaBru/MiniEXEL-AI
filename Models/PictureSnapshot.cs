using System.Windows.Media.Imaging;

namespace MiniEXEL.Models
{
    // Approximate placement of an embedded picture: anchored to the top-left
    // cell it was inserted at (as Excel stores it), rendered at a fixed
    // row-height/column-width grid so its position lines up closely, not pixel-exact.
    public class PictureSnapshot
    {
        public BitmapImage Image { get; set; } = null!;
        public int AnchorRowIndex { get; set; } // 0-based
        public int AnchorColIndex { get; set; } // 0-based
        public double WidthPx { get; set; }
        public double HeightPx { get; set; }
    }
}
