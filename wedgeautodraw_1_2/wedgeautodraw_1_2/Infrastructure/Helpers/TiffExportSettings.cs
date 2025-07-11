using DocumentFormat.OpenXml.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

using System;
using System.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Tiff;
using SixLabors.ImageSharp.Formats.Tiff.Constants;
using SixLabors.ImageSharp.Metadata;
using ImageMagick;
using ImageMagick.Formats;


namespace wedgeautodraw_1_2.Infrastructure.Helpers;

public class TiffExportSettings
{
    private readonly SldWorks _swApp;
    private readonly ModelDoc2 _model;

    public TiffExportSettings(SldWorks swApp)
    {
        _swApp = swApp;
    }

    public bool SetTiffExportSettings(int dpi = 100)
    {
        // 640×480 pixels at 100 DPI → 0.16256 m × 0.12192 m
        double widthMeters = 640 / 100.0 * 0.0254;  // = 0.16256
        double heightMeters = 480 / 100.0 * 0.0254; // = 0.12192

        try
        {
            Logger.Info($"Setting TIFF export settings: DPI={dpi}, Width={widthMeters:F5}m, Height={heightMeters:F5}m");
            _swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, 1);

            bool dpiSet = _swApp.SetUserPreferenceIntegerValue(
                (int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, dpi);

            bool widthSet = _swApp.SetUserPreferenceDoubleValue(
                (int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, widthMeters);

            bool heightSet = _swApp.SetUserPreferenceDoubleValue(
                (int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, heightMeters);

            if (dpiSet && widthSet && heightSet)
            {
                Logger.Success("TIFF export settings applied successfully.");
                return true;
            }

            Logger.Warn("Some TIFF export preferences failed to apply.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to set TIFF export preferences: {ex.Message}");
            return false;
        }
    }

    public void PrintTiffExportSettings()
    {
        try
        {
            int dpi = _swApp.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintDPI);
            double width = _swApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth);
            double height = _swApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight);

            Logger.Info($"Current TIFF Export Settings:");
            Logger.Info($"  DPI           : {dpi}");
            Logger.Info($"  Paper Width   : {width:F4} meters ({width * 39.3701:F2} inches)");
            Logger.Info($"  Paper Height  : {height:F4} meters ({height * 39.3701:F2} inches)");

            double widthPx = width * dpi / 0.0254;
            double heightPx = height * dpi / 0.0254;

            Logger.Info($"  Pixel Size    : {Math.Round(widthPx)} x {Math.Round(heightPx)} pixels");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to read TIFF export preferences: {ex.Message}");
        }
    }

    public bool RunSolidWorksMacro(SldWorks swApp, string macroPath)
    {
        int errorCode;
        bool result = swApp.RunMacro2(
            macroPath,
            "Macro1", // Your module name in the .swp macro
            "main",   // Your subroutine name in the macro
            (int)swRunMacroOption_e.swRunMacroUnloadAfterRun,
            out errorCode
        );

        if (result)
            Logger.Success("Macro executed successfully.");
        else
            Logger.Error($"Macro execution failed with error code: {errorCode}");

        return result;
    }

    public void ResizeTiffTo640x480(string inputFile, string outputFile)
    {
        try
        {
            using (var image = System.Drawing.Image.FromFile(inputFile))
            {
                int srcWidth = image.Width;
                int srcHeight = image.Height;

                int targetWidth = 640;
                int targetHeight = 480;

                using (var bmp = new System.Drawing.Bitmap(targetWidth, targetHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    bmp.SetResolution(300, 300); // Set high DPI for quality

                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                        var destRect = new System.Drawing.Rectangle(0, 0, targetWidth, targetHeight);
                        var srcRect = new System.Drawing.Rectangle(0, 0, srcWidth, srcHeight);

                        g.DrawImage(image, destRect, srcRect, System.Drawing.GraphicsUnit.Pixel);
                    }

                    var tiffCodec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(codec => codec.FormatID == System.Drawing.Imaging.ImageFormat.Tiff.Guid);

                    var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                    encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.Compression,
                        (long)System.Drawing.Imaging.EncoderValue.CompressionLZW
                    );

                    bmp.Save(outputFile, tiffCodec, encoderParams);
                    Logger.Success($"TIFF resized and saved to: {outputFile}");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to resize TIFF: {ex.Message}");
        }
    }

    public void ResizeTiffWithImageSharp(string inputFile, string outputFile)
    {
        try
        {
            using (var image = SixLabors.ImageSharp.Image.Load<Rgba32>(inputFile))
            {
                image.Mutate(x => x.Resize(new ResizeOptions
                {
                    Size = new SixLabors.ImageSharp.Size(640, 480),
                    Mode = ResizeMode.Max,
                    Sampler = KnownResamplers.Lanczos3
                }));

                var encoder = new TiffEncoder
                {
                    Compression = TiffCompression.PackBits,
                    BitsPerPixel = TiffBitsPerPixel.Bit64
                };

                image.Save(outputFile, encoder);

                Logger.Success($"[ImageSharp] TIFF resized and saved to: {outputFile}");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[ImageSharp] Failed to resize TIFF: {ex.Message}");
        }
    }

    public void ResizeTiffWithGdiPreservingQuality(string inputPath, string outputPath, int maxHeight = 480)
    {
        try
        {
            using (var image = System.Drawing.Image.FromFile(inputPath))
            {
                var resized = ResizePreservingAspectRatio(image, maxHeight);

                // Use LZW compression to preserve quality
                var codec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                    .FirstOrDefault(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Tiff.Guid);

                var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Compression,
                    (long)System.Drawing.Imaging.EncoderValue.CompressionLZW
                );

                resized.Save(outputPath, codec, encoderParams);
                resized.Dispose();

                Logger.Success($"[GDI+] TIFF resized and saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[GDI+] Failed to resize TIFF: {ex.Message}");
        }
    }

    private System.Drawing.Bitmap ResizePreservingAspectRatio(System.Drawing.Image image, int maxHeight)
    {
        double scale = (double)image.Height / maxHeight;
        int targetWidth = (int)(image.Width / scale);
        int targetHeight = maxHeight;

        var destRect = new System.Drawing.Rectangle(0, 0, targetWidth, targetHeight);
        var destImage = new System.Drawing.Bitmap(targetWidth, targetHeight);
        destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

        using (var graphics = System.Drawing.Graphics.FromImage(destImage))
        {
            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
            graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

            using (var wrapMode = new System.Drawing.Imaging.ImageAttributes())
            {
                wrapMode.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, System.Drawing.GraphicsUnit.Pixel, wrapMode);
            }
        }

        return destImage;
    }
   public void ResizeImageTo640x480WithDpi(string inputPath, string outputPath)
    {
        try
        {
            using (var image = System.Drawing.Image.FromFile(inputPath))
            {
                const int targetWidth = 640;
                const int targetHeight = 480;
                const int targetDpi = 100;

                using (var canvas = new System.Drawing.Bitmap(targetWidth, targetHeight, System.Drawing.Imaging.PixelFormat.Format32bppPArgb))
                {
                    canvas.SetResolution(targetDpi, targetDpi);

                    using (var graphics = System.Drawing.Graphics.FromImage(canvas))
                    {
                        graphics.Clear(System.Drawing.Color.White); // Optional: white background
                        graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                        float scaleX = (float)targetWidth / image.Width;
                        float scaleY = (float)targetHeight / image.Height;
                        float scale = Math.Min(scaleX, scaleY);

                        int drawWidth = (int)(image.Width * scale);
                        int drawHeight = (int)(image.Height * scale);

                        int offsetX = (targetWidth - drawWidth) / 2;
                        int offsetY = (targetHeight - drawHeight) / 2;

                        var destRect = new System.Drawing.Rectangle(offsetX, offsetY, drawWidth, drawHeight);

                        using (var wrapMode = new System.Drawing.Imaging.ImageAttributes())
                        {
                            wrapMode.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                            graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, System.Drawing.GraphicsUnit.Pixel, wrapMode);
                        }
                    }

                    var tiffCodec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Tiff.Guid);

                    var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                    encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.Compression,
                        (long)System.Drawing.Imaging.EncoderValue.CompressionLZW
                    );

                    canvas.Save(outputPath, tiffCodec, encoderParams);
                    Logger.Success($"[GDI+] High-quality TIFF saved to: {outputPath}");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[GDI+] ResizeImageTo640x480WithDpi failed: {ex.Message}");
        }
    }
    // the working one
    public void ResizeImageSharpHighQuality(string inputPath, string outputPath, bool applySharpen = true)
    {
        try
        {
            using (var image = SixLabors.ImageSharp.Image.Load<Rgba32>(inputPath))
            {
                const int targetWidth = 1280;
                const int targetHeight = 1024;
                const int dpi = 300;

                image.Mutate(ctx =>
                {
                    ctx.Resize(new ResizeOptions
                    {
                        Size = new SixLabors.ImageSharp.Size(targetWidth, targetHeight),
                        Position = AnchorPositionMode.Center,
                        Sampler = KnownResamplers.Lanczos3,
                        Mode = ResizeMode.Pad
                    });

                    if (applySharpen)
                    {
                        ctx.GaussianSharpen(1.0f);
                    }
                });

                image.Metadata.VerticalResolution = dpi;
                image.Metadata.HorizontalResolution = dpi;

                // Flatten transparency to white to prevent dark colors
                using (var flattened = image.Clone(x => x.BackgroundColor(SixLabors.ImageSharp.Color.White)))
                {
                    var encoder = new TiffEncoder
                    {
                        Compression = TiffCompression.Lzw,
                        BitsPerPixel = TiffBitsPerPixel.Bit64
                    };

                    flattened.Save(outputPath, encoder);
                }

                Logger.Success($"[ImageSharp] TIFF saved to: {outputPath} (640x480 @ {dpi} DPI)");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[ImageSharp] Resize failed: {ex.Message}");
        }
    }

    public void ResizeTiffUsingTiffLibrary(string inputPath, string outputPath)
    {
        try
        {
            using (var image = SixLabors.ImageSharp.Image.Load<SixLabors.ImageSharp.PixelFormats.Rgba32>(inputPath))
            {
                // Resize image to exactly 640x480 using high-quality Lanczos resampling
                image.Mutate(x => x.Resize(new ResizeOptions
                {
                    Size = new SixLabors.ImageSharp.Size(640, 480),
                    Mode = ResizeMode.Stretch,
                    Sampler = KnownResamplers.Lanczos3
                }));

                // Set 100 DPI resolution
                image.Metadata.HorizontalResolution = 100;
                image.Metadata.VerticalResolution = 100;
                image.Metadata.ResolutionUnits = PixelResolutionUnit.PixelsPerInch;

                // Save as TIFF using TiffLibrary-backed encoder
                var encoder = new TiffEncoder
                {
                    Compression = TiffCompression.Lzw,
                    BitsPerPixel = TiffBitsPerPixel.Bit24
                };

                image.Save(outputPath, encoder);
                Logger.Success($"[TiffLibrary] Resized TIFF saved to: {outputPath} (640x480 @ 100 DPI)");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[TiffLibrary] Resize failed: {ex.Message}");
        }
    }
    public void ResizeAndCenterImageTo480x640(string inputFile, string outputFile)
    {
        try
        {
            using (var image = System.Drawing.Image.FromFile(inputFile))
            {
                int srcWidth = image.Width;
                int srcHeight = image.Height;

                int targetWidth = 640;
                int targetHeight = 480;

                // Calculate aspect ratio
                float srcAspect = (float)srcWidth / srcHeight;
                float targetAspect = (float)targetWidth / targetHeight;

                int resizeWidth = targetWidth;
                int resizeHeight = targetHeight;

                if (srcAspect > targetAspect)
                {
                    resizeHeight = (int)(targetWidth / srcAspect);
                }
                else
                {
                    resizeWidth = (int)(targetHeight * srcAspect);
                }

                using (var bmp = new System.Drawing.Bitmap(targetWidth, targetHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        // White background
                        g.Clear(System.Drawing.Color.White);

                        // High quality settings
                        g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                        // Center the image
                        int offsetX = (targetWidth - resizeWidth) / 2;
                        int offsetY = (targetHeight - resizeHeight) / 2;
                        var destRect = new System.Drawing.Rectangle(offsetX, offsetY, resizeWidth, resizeHeight);

                        g.DrawImage(image, destRect, 0, 0, srcWidth, srcHeight, System.Drawing.GraphicsUnit.Pixel);
                    }

                    // Save TIFF with LZW compression
                    var tiffCodec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(codec => codec.FormatID == System.Drawing.Imaging.ImageFormat.Tiff.Guid);

                    var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                    encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.Compression,
                        (long)System.Drawing.Imaging.EncoderValue.CompressionLZW);

                    bmp.Save(outputFile, tiffCodec, encoderParams);
                    Logger.Success($"[GDI+] Resized TIFF saved to: {outputFile} (480x640 centered)");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[GDI+] ResizeAndCenterImageTo480x640 failed: {ex.Message}");
        }
    }
    public void ResizeTiffWithMagickNet_HighQuality(string inputPath, string outputPath)
    {
        try
        {
            using (var image = new MagickImage(inputPath))
            {
                // Set 100 DPI
                image.Density = new Density(100, 100);

                // Resize with aspect ratio preserved and padding
                var geometry = new MagickGeometry(640, 480)
                {
                    IgnoreAspectRatio = false
                };
                image.Resize(geometry);

                // Center on a white canvas if image doesn't exactly fill 640x480
                image.Extent(640, 480, Gravity.Center, MagickColors.White);

                // Ensure alpha is removed for clean TIFF background
                image.Alpha(AlphaOption.Remove);

                // Set color type to true color for rich output
                image.ColorType = ImageMagick.ColorType.TrueColor;

                // Set TIFF format explicitly
                image.Format = MagickFormat.Tiff;

                // Optional: Strip metadata to avoid bloat
                image.Strip();

                // Save to output file
                image.Write(outputPath);

                Logger.Success($"[Magick.NET] High-quality TIFF saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"[Magick.NET] Resize failed: {ex.Message}");
        }
    }

}
