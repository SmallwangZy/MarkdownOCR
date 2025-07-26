using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MarkdownOCR
{
    /// <summary>
    /// 轻量化截图工具 - 启动后立即进入截图模式，支持鼠标拖拽选择区域
    /// 松开鼠标后保存截图到当前目录并退出程序
    /// </summary>
    public partial class ScreenCaptureForm : Form
    {
        private bool _isSelecting = false;
        private Point _startPoint;
        private Point _endPoint;
        private Rectangle _selectionRectangle;
        private Bitmap? _screenBitmap;
        
        public ScreenCaptureForm()
        {
            InitializeComponent();
            CaptureScreen();
        }
        
        /// <summary>
        /// 初始化窗体组件和属性
        /// </summary>
        private void InitializeComponent()
        {
            // 设置窗体为全屏透明覆盖层
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.BackColor = Color.Black;
            this.Opacity = 0.3; // 半透明背景
            this.Cursor = Cursors.Cross;
            this.DoubleBuffered = true; // 减少闪烁
            
            // 绑定鼠标事件
            this.MouseDown += OnMouseDown;
            this.MouseMove += OnMouseMove;
            this.MouseUp += OnMouseUp;
            this.Paint += OnPaint;
            this.KeyDown += OnKeyDown;
        }
        
        /// <summary>
        /// 捕获整个屏幕到内存
        /// </summary>
        private void CaptureScreen()
        {
            try
            {
                // 获取主屏幕尺寸
                Rectangle screenBounds = Screen.PrimaryScreen?.Bounds ?? Screen.AllScreens[0].Bounds;
                _screenBitmap = new Bitmap(screenBounds.Width, screenBounds.Height, PixelFormat.Format32bppArgb);
                
                // 使用Graphics.CopyFromScreen捕获屏幕
                using (Graphics graphics = Graphics.FromImage(_screenBitmap))
                {
                    graphics.CopyFromScreen(screenBounds.Left, screenBounds.Top, 0, 0, screenBounds.Size);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"屏幕捕获失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
        
        /// <summary>
        /// 鼠标按下事件 - 开始选择区域
        /// </summary>
        private void OnMouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _isSelecting = true;
                _startPoint = e.Location;
                _endPoint = e.Location;
                _selectionRectangle = new Rectangle();
            }
            else if (e.Button == MouseButtons.Right)
            {
                // 右键取消截图
                Application.Exit();
            }
        }
        
        /// <summary>
        /// 鼠标移动事件 - 更新选择区域
        /// </summary>
        private void OnMouseMove(object? sender, MouseEventArgs e)
        {
            if (_isSelecting)
            {
                _endPoint = e.Location;
                
                // 计算选择矩形（支持任意方向拖拽）
                int x = Math.Min(_startPoint.X, _endPoint.X);
                int y = Math.Min(_startPoint.Y, _endPoint.Y);
                int width = Math.Abs(_endPoint.X - _startPoint.X);
                int height = Math.Abs(_endPoint.Y - _startPoint.Y);
                
                _selectionRectangle = new Rectangle(x, y, width, height);
                
                // 重绘窗体以显示选择框
                this.Invalidate();
            }
        }
        
        /// <summary>
        /// 鼠标释放事件 - 完成选择并保存截图
        /// </summary>
        private void OnMouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && _isSelecting)
            {
                _isSelecting = false;
                
                // 检查选择区域是否有效
                if (_selectionRectangle.Width > 5 && _selectionRectangle.Height > 5)
                {
                    SaveSelectedArea();
                }
                else
                {
                    // 选择区域太小，取消截图
                    Application.Exit();
                }
            }
        }
        
        /// <summary>
        /// 处理选中区域的截图并进行OCR识别
        /// </summary>
        private async void SaveSelectedArea()
        {
            if (_screenBitmap == null) return;
            
            // 保存截图区域坐标，用于悬浮窗定位
            Rectangle captureRect = _selectionRectangle;
            
            try
            {
                // 立即隐藏截图界面，恢复正常屏幕状态
                this.Hide();
                
                // 从完整屏幕截图中裁剪选中区域
                using (Bitmap croppedBitmap = new Bitmap(_selectionRectangle.Width, _selectionRectangle.Height))
                {
                    using (Graphics graphics = Graphics.FromImage(croppedBitmap))
                    {
                        graphics.DrawImage(_screenBitmap, 
                            new Rectangle(0, 0, _selectionRectangle.Width, _selectionRectangle.Height),
                            _selectionRectangle, 
                            GraphicsUnit.Pixel);
                    }
                    
                    // 直接进行图像预处理和OCR识别，不保存到本地文件
                    await PerformOCRRecognition(croppedBitmap, captureRect);
                }
            }
            catch (Exception ex)
            {
                // 显示错误信息悬浮窗
                FloatingWindowManager.ShowError($"截图处理失败: {ex.Message}", captureRect);
            }
        }
        
        /// <summary>
        /// 执行OCR文字识别
        /// </summary>
        /// <param name="bitmap">要识别的图像</param>
        /// <param name="captureRect">截图区域坐标</param>
        private async Task PerformOCRRecognition(Bitmap bitmap, Rectangle captureRect)
        {
            try
            {
                // 1. 图像预处理并获取Base64编码
                string base64Image = ImageProcessor.ProcessScreenshotForOCR(bitmap);
                
                // 2. 向Ollama发送OCR请求并获取处理时间
                var (recognizedText, processingTime) = await OllamaClient.RecognizeTextAsync(base64Image);
                
                // 3. 显示识别结果悬浮窗（包含处理时间）
                FloatingWindowManager.ShowOCRResult(recognizedText, captureRect, processingTime);
            }
            catch (Exception ex)
            {
                // 显示错误信息悬浮窗
                FloatingWindowManager.ShowError($"OCR识别失败: {ex.Message}", captureRect);
            }
        }
        
        /// <summary>
        /// 绘制选择框
        /// </summary>
        private void OnPaint(object? sender, PaintEventArgs e)
        {
            if (_isSelecting && _selectionRectangle.Width > 0 && _selectionRectangle.Height > 0)
            {
                // 绘制白色虚线选择框边框
                using (Pen pen = new Pen(Color.White, 2))
                {
                    pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                    e.Graphics.DrawRectangle(pen, _selectionRectangle);
                }
                
                // 在选择框内部绘制半透明白色背景，提高可见性
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(50, Color.White)))
                {
                    e.Graphics.FillRectangle(brush, _selectionRectangle);
                }
            }
        }
        
        /// <summary>
        /// 键盘事件 - ESC键取消截图
        /// </summary>
        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Application.Exit();
            }
        }
        
        /// <summary>
        /// 释放资源
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _screenBitmap?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
    
    /// <summary>
    /// 程序入口点
    /// </summary>
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            try
            {
                // 启用应用程序视觉样式
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                // 启动截图窗体
                using (var captureForm = new ScreenCaptureForm())
                {
                    Application.Run(captureForm);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"程序启动失败: {ex.Message}\n\n详细错误: {ex}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}