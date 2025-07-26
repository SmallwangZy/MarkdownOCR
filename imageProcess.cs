using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace MarkdownOCR
{
    /// <summary>
    /// 图像处理类，负责图像预处理和Base64编码
    /// </summary>
    public static class ImageProcessor
    {
        /// <summary>
        /// 将Bitmap图像转换为Base64编码字符串
        /// </summary>
        /// <param name="bitmap">要转换的Bitmap图像</param>
        /// <returns>Base64编码的图像字符串</returns>
        public static string ConvertToBase64(Bitmap bitmap)
        {
            if (bitmap == null)
                throw new ArgumentNullException(nameof(bitmap));

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    // 将图像保存为PNG格式到内存流
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    
                    // 转换为Base64字符串
                    byte[] imageBytes = memoryStream.ToArray();
                    return Convert.ToBase64String(imageBytes);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"图像转换Base64失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 对图像进行预处理以提高OCR识别准确率
        /// </summary>
        /// <param name="originalBitmap">原始图像</param>
        /// <returns>预处理后的图像</returns>
        public static Bitmap PreprocessImage(Bitmap originalBitmap)
        {
            if (originalBitmap == null)
                throw new ArgumentNullException(nameof(originalBitmap));

            try
            {
                // 创建预处理后的图像副本
                Bitmap processedBitmap = new Bitmap(originalBitmap.Width, originalBitmap.Height, PixelFormat.Format24bppRgb);
                
                using (Graphics graphics = Graphics.FromImage(processedBitmap))
                {
                    // 设置高质量渲染
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    
                    // 绘制原始图像到新图像
                    graphics.DrawImage(originalBitmap, 0, 0, originalBitmap.Width, originalBitmap.Height);
                }

                // 应用图像增强处理
                EnhanceImageForOCR(processedBitmap);
                
                return processedBitmap;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"图像预处理失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 增强图像以提高OCR识别效果
        /// </summary>
        /// <param name="bitmap">要增强的图像</param>
        private static void EnhanceImageForOCR(Bitmap bitmap)
        {
            // 锁定位图数据以进行像素级操作
            BitmapData bitmapData = bitmap.LockBits(
                new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                ImageLockMode.ReadWrite,
                PixelFormat.Format24bppRgb);

            try
            {
                unsafe
                {
                    byte* ptr = (byte*)bitmapData.Scan0;
                    int bytesPerPixel = 3; // RGB格式
                    int stride = bitmapData.Stride;

                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            int offset = y * stride + x * bytesPerPixel;
                            
                            // 获取RGB值
                            byte blue = ptr[offset];
                            byte green = ptr[offset + 1];
                            byte red = ptr[offset + 2];
                            
                            // 计算灰度值（加权平均）
                            byte gray = (byte)(0.299 * red + 0.587 * green + 0.114 * blue);
                            
                            // 应用对比度增强
                            gray = EnhanceContrast(gray);
                            
                            // 设置增强后的像素值
                            ptr[offset] = gray;     // Blue
                            ptr[offset + 1] = gray; // Green
                            ptr[offset + 2] = gray; // Red
                        }
                    }
                }
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
        }

        /// <summary>
        /// 增强像素对比度
        /// </summary>
        /// <param name="value">原始像素值</param>
        /// <returns>增强后的像素值</returns>
        private static byte EnhanceContrast(byte value)
        {
            // 应用S曲线对比度增强
            double normalized = value / 255.0;
            double enhanced = Math.Pow(normalized, 0.8); // 轻微增强对比度
            
            // 确保值在有效范围内
            enhanced = Math.Max(0, Math.Min(1, enhanced));
            
            return (byte)(enhanced * 255);
        }

        /// <summary>
        /// 处理截图图像并返回Base64编码
        /// </summary>
        /// <param name="originalBitmap">原始截图</param>
        /// <returns>预处理后的Base64编码图像</returns>
        public static string ProcessScreenshotForOCR(Bitmap originalBitmap)
        {
            if (originalBitmap == null)
                throw new ArgumentNullException(nameof(originalBitmap));

            try
            {
                // 预处理图像
                using (Bitmap processedBitmap = PreprocessImage(originalBitmap))
                {
                    // 转换为Base64
                    return ConvertToBase64(processedBitmap);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"截图处理失败: {ex.Message}", ex);
            }
        }
    }
}