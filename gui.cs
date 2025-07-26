using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Markdig;

namespace MarkdownOCR
{
    /// <summary>
    /// OCR结果悬浮窗管理器
    /// </summary>
    public partial class OCRResultWindow : Form
    {
        private readonly string _ocrResult;
        private readonly Rectangle _sourceRect;
        private WebBrowser? _resultWebBrowser;
        private Button? _copyButton;
        private Button? _closeButton;
        private Label? _timeLabel;
        private string _originalMarkdownContent = string.Empty;
        private double _processingTimeSeconds = 0.0;
        
        // 窗口尺寸常量
        private const int MinWindowWidth = 450;
        private const int MinWindowHeight = 250;
        private const int MaxWindowWidth = 800;
        private const int MaxWindowHeight = 600;
        private const int ButtonWidth = 35;
        private const int ButtonHeight = 35;
        private new const int Padding = 12;
        private const int ButtonSpacing = 15;
        private const int TimeLabelHeight = 25;
        
        public OCRResultWindow(string ocrResult, Rectangle sourceRect) : this(ocrResult, sourceRect, 0.0)
        {
        }
        
        public OCRResultWindow(string ocrResult, Rectangle sourceRect, double processingTimeSeconds)
        {
            _ocrResult = ocrResult ?? "未识别到任何文字内容";
            _sourceRect = sourceRect;
            _processingTimeSeconds = processingTimeSeconds;
            _originalMarkdownContent = _ocrResult; // 保存原始Markdown内容
            InitializeComponent();
            SetWindowPosition();
        }
        
        /// <summary>
        /// 初始化悬浮窗组件
        /// </summary>
        private void InitializeComponent()
        {
            // 保存原始Markdown内容
            _originalMarkdownContent = _ocrResult;
            
            // 计算自适应窗口大小
            var windowSize = CalculateOptimalWindowSize(_ocrResult);
            
            // 窗体基本设置
            this.Text = "OCR识别结果";
            this.Size = windowSize;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true; // 置顶显示
            this.StartPosition = FormStartPosition.Manual;
            this.BackColor = Color.White;
            this.Font = new Font("微软雅黑", 9F);
            
            // 计算文本框和按钮位置
            int textBoxWidth = windowSize.Width - ButtonWidth - Padding * 4;
            int textBoxHeight = windowSize.Height - Padding * 2;
            
            // 创建时间标签
            _timeLabel = new Label
            {
                Location = new Point(Padding, Padding),
                Size = new Size(textBoxWidth, TimeLabelHeight),
                Text = _processingTimeSeconds > 0 ? $"消耗时间: {_processingTimeSeconds:F2} s" : "",
                Font = new Font("Microsoft YaHei", 9, FontStyle.Regular),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleLeft,
                Visible = _processingTimeSeconds > 0
            };
            
            // 计算WebBrowser的位置和大小
            int webBrowserTop = _timeLabel.Visible ? _timeLabel.Bottom + 5 : Padding;
            int webBrowserHeight = textBoxHeight - (_timeLabel.Visible ? TimeLabelHeight + 5 : 0);
            
            // 创建WebBrowser控件（左侧占主要空间）
            _resultWebBrowser = new WebBrowser
            {
                Location = new Point(Padding, webBrowserTop),
                Size = new Size(textBoxWidth, webBrowserHeight),
                ScrollBarsEnabled = true,
                IsWebBrowserContextMenuEnabled = false,
                WebBrowserShortcutsEnabled = false,
                AllowNavigation = false
            };
            
            // 渲染Markdown内容
            RenderMarkdownContent(_ocrResult);
            
            // 计算按钮位置（右侧垂直居中）
            int buttonX = textBoxWidth + Padding * 2;
            int centerY = windowSize.Height / 2;
            int copyButtonY = centerY - ButtonHeight - ButtonSpacing / 2;
            int closeButtonY = centerY + ButtonSpacing / 2;
            
            // 创建复制按钮（右侧上方）
            _copyButton = new Button
            {
                Location = new Point(buttonX, copyButtonY),
                Size = new Size(ButtonWidth, ButtonHeight),
                Text = "📋",
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(0, 122, 255),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Emoji", 14F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _copyButton.FlatAppearance.BorderSize = 0;
            _copyButton.FlatAppearance.BorderColor = Color.FromArgb(0, 122, 255);
            _copyButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 220);
            _copyButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 80, 180);
            _copyButton.Click += OnCopyButtonClick;
            
            // 创建关闭按钮（右侧下方）
            _closeButton = new Button
            {
                Location = new Point(buttonX, closeButtonY),
                Size = new Size(ButtonWidth, ButtonHeight),
                Text = "✕",
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(255, 59, 48),
                ForeColor = Color.White,
                Font = new Font("Arial", 16F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _closeButton.FlatAppearance.BorderSize = 0;
            _closeButton.FlatAppearance.BorderColor = Color.FromArgb(255, 59, 48);
            _closeButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 50, 40);
            _closeButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(180, 40, 30);
            _closeButton.Click += OnCloseButtonClick;
            
            // 添加控件到窗体
            if (_timeLabel.Visible)
                this.Controls.Add(_timeLabel);
            this.Controls.Add(_resultWebBrowser);
            this.Controls.Add(_copyButton);
            this.Controls.Add(_closeButton);
            
            // 设置默认按钮和取消按钮
            this.AcceptButton = _copyButton;
            this.CancelButton = _closeButton;
            
            // 绑定键盘事件
            this.KeyPreview = true;
            this.KeyDown += OnKeyDown;
        }
        
        /// <summary>
        /// 计算最佳窗口大小
        /// </summary>
        /// <param name="content">文本内容</param>
        /// <returns>最佳窗口尺寸</returns>
        private Size CalculateOptimalWindowSize(string content)
        {
            if (string.IsNullOrEmpty(content))
                return new Size(MinWindowWidth, MinWindowHeight);
            
            // 获取屏幕工作区域
            Rectangle workingArea = Screen.PrimaryScreen?.WorkingArea ?? Screen.AllScreens[0].WorkingArea;
            
            // 计算文本行数和最长行字符数
            string[] lines = content.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            int lineCount = Math.Max(lines.Length, 1);
            int maxLineLength = lines.Length > 0 ? lines.Max(line => line.Length) : 0;
            
            // 基于字符数估算宽度（每个字符约8像素，中文字符约16像素）
            int estimatedWidth = (int)(maxLineLength * 12) + ButtonWidth + Padding * 4;
            
            // 更精确的高度计算：考虑文本换行和实际渲染
            // 计算在给定宽度下的实际行数（考虑自动换行）
            int contentWidth = estimatedWidth - ButtonWidth - Padding * 4;
            int actualLineCount = 0;
            
            foreach (string line in lines)
            {
                if (string.IsNullOrEmpty(line))
                {
                    actualLineCount += 1;
                }
                else
                {
                    // 估算该行在给定宽度下会占用多少行
                    int linePixelWidth = line.Length * 12; // 每个字符约12像素
                    int linesForThisContent = Math.Max(1, (int)Math.Ceiling((double)linePixelWidth / contentWidth));
                    actualLineCount += linesForThisContent;
                }
            }
            
            // 基于实际行数估算高度（每行约22像素），考虑时间标签的高度
            int timeLabelSpace = _processingTimeSeconds > 0 ? TimeLabelHeight + 5 : 0;
            int estimatedHeight = actualLineCount * 22 + Padding * 3 + timeLabelSpace + 50; // 额外50像素缓冲
            
            // 参考截图区域尺寸进行智能调整
            // 悬浮窗应该能够完整显示文本，同时参考截图区域大小
            
            // 宽度：优先保证文本显示，但不超过截图区域的1.8倍或屏幕的80%
            int maxWidthFromCapture = Math.Max(_sourceRect.Width, (int)(_sourceRect.Width * 1.8));
            int maxWidthFromScreen = (int)(workingArea.Width * 0.8);
            int maxAllowedWidth = Math.Min(maxWidthFromCapture, maxWidthFromScreen);
            
            // 高度：优先保证文本完整显示，参考截图区域但允许更大的灵活性
            int maxHeightFromCapture = Math.Max(_sourceRect.Height, (int)(_sourceRect.Height * 2.0));
            int maxHeightFromScreen = (int)(workingArea.Height * 0.8);
            int maxAllowedHeight = Math.Min(maxHeightFromCapture, maxHeightFromScreen);
            
            // 应用所有限制条件，优先保证文本完整显示
            int finalWidth = Math.Max(MinWindowWidth, 
                Math.Min(Math.Min(MaxWindowWidth, Math.Max(estimatedWidth, maxAllowedWidth)), maxWidthFromScreen));
            int finalHeight = Math.Max(MinWindowHeight, 
                Math.Min(Math.Min(MaxWindowHeight, Math.Max(estimatedHeight, maxAllowedHeight)), maxHeightFromScreen));
            
            // 如果估算的尺寸超过了基于截图的限制，优先保证文本显示
            if (estimatedWidth > maxAllowedWidth)
            {
                finalWidth = Math.Min(estimatedWidth, maxWidthFromScreen);
            }
            if (estimatedHeight > maxAllowedHeight)
            {
                finalHeight = Math.Min(estimatedHeight, maxHeightFromScreen);
            }
            
            return new Size(finalWidth, finalHeight);
        }
        
        /// <summary>
        /// 渲染Markdown内容到WebBrowser
        /// </summary>
        /// <param name="markdownContent">Markdown格式的内容</param>
        private void RenderMarkdownContent(string markdownContent)
        {
            if (_resultWebBrowser == null || string.IsNullOrEmpty(markdownContent))
                return;
            
            try
            {
                // 使用Markdig将Markdown转换为HTML
                var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
                string html = Markdown.ToHtml(markdownContent, pipeline);
                
                // 创建完整的HTML文档
                string fullHtml = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"">
    <style>
        body {{
            font-family: '微软雅黑', 'Microsoft YaHei', Arial, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: #333;
            margin: 10px;
            background-color: #fff;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: #2c3e50;
            margin-top: 20px;
            margin-bottom: 10px;
        }}
        h1 {{ font-size: 24px; }}
        h2 {{ font-size: 20px; }}
        h3 {{ font-size: 16px; }}
        code {{
            background-color: #f8f8f8;
            border: 1px solid #e1e1e8;
            border-radius: 3px;
            padding: 2px 4px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            color: #c7254e;
        }}
        pre {{
            background-color: #f8f8f8;
            border: 1px solid #e1e1e8;
            border-radius: 4px;
            padding: 10px;
            overflow-x: auto;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
        }}
        ul, ol {{
            padding-left: 20px;
        }}
        li {{
            margin-bottom: 5px;
        }}
        blockquote {{
            border-left: 4px solid #ddd;
            margin: 0;
            padding-left: 16px;
            color: #666;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #f2f2f2;
        }}
    </style>
</head>
<body>
{html}
</body>
</html>";
                
                _resultWebBrowser.DocumentText = fullHtml;
            }
            catch (Exception)
            {
                // 如果Markdown渲染失败，显示原始文本
                string errorHtml = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"">
    <style>
        body {{
            font-family: '微软雅黑', Arial, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: #333;
            margin: 10px;
            white-space: pre-wrap;
        }}
    </style>
</head>
<body>
{markdownContent.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")}
</body>
</html>";
                
                _resultWebBrowser.DocumentText = errorHtml;
            }
        }
        
        /// <summary>
        /// 复制原始Markdown内容到剪贴板
        /// </summary>
        private void CopyToClipboard()
        {
            try
            {
                if (!string.IsNullOrEmpty(_originalMarkdownContent))
                {
                    Clipboard.SetText(_originalMarkdownContent);
                    
                    // 视觉反馈：按钮变绿并显示✓
                    if (_copyButton != null)
                    {
                        _copyButton.BackColor = Color.FromArgb(40, 167, 69);
                        _copyButton.Text = "✓";
                        
                        // 1秒后恢复原状
                        var timer = new System.Windows.Forms.Timer { Interval = 1000 };
                        timer.Tick += (s, e) =>
                        {
                            if (_copyButton != null)
                            {
                                _copyButton.BackColor = Color.FromArgb(0, 123, 255);
                                _copyButton.Text = "📋";
                            }
                            timer.Stop();
                            timer.Dispose();
                        };
                        timer.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        

        
        /// <summary>
        /// 设置悬浮窗位置（在截图区域附近显示）
        /// </summary>
        private void SetWindowPosition()
        {
            // 获取屏幕工作区域
            Rectangle workingArea = Screen.PrimaryScreen?.WorkingArea ?? Screen.AllScreens[0].WorkingArea;
            
            // 计算悬浮窗的理想位置（截图区域右下角）
            int x = _sourceRect.Right + 10;
            int y = _sourceRect.Bottom + 10;
            
            // 确保悬浮窗不会超出屏幕边界
            if (x + this.Width > workingArea.Right)
            {
                x = _sourceRect.Left - this.Width - 10;
                if (x < workingArea.Left)
                {
                    x = workingArea.Right - this.Width - 10;
                }
            }
            
            if (y + this.Height > workingArea.Bottom)
            {
                y = _sourceRect.Top - this.Height - 10;
                if (y < workingArea.Top)
                {
                    y = workingArea.Bottom - this.Height - 10;
                }
            }
            
            // 设置窗体位置
            this.Location = new Point(Math.Max(x, workingArea.Left), Math.Max(y, workingArea.Top));
        }
        
        /// <summary>
        /// 复制按钮点击事件
        /// </summary>
        private void OnCopyButtonClick(object? sender, EventArgs e)
        {
            CopyToClipboard();
        }
        

        
        /// <summary>
        /// 关闭按钮点击事件
        /// </summary>
        private void OnCloseButtonClick(object? sender, EventArgs e)
        {
            this.Close();
            Application.Exit(); // 关闭整个应用程序
        }
        
        /// <summary>
        /// 键盘事件处理
        /// </summary>
        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
                Application.Exit();
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                OnCopyButtonClick(sender, e);
            }
        }
        
        /// <summary>
        /// 窗体关闭事件
        /// </summary>
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            Application.Exit(); // 确保程序完全退出
        }
    }
    
    /// <summary>
    /// 悬浮窗管理器静态类
    /// </summary>
    public static class FloatingWindowManager
    {
        /// <summary>
        /// 显示OCR结果悬浮窗
        /// </summary>
        /// <param name="ocrResult">OCR识别结果</param>
        /// <param name="sourceRect">截图区域位置</param>
        public static void ShowOCRResult(string ocrResult, Rectangle sourceRect)
        {
            ShowOCRResult(ocrResult, sourceRect, 0.0);
        }
        
        /// <summary>
        /// 显示OCR结果悬浮窗（带处理时间）
        /// </summary>
        /// <param name="ocrResult">OCR识别结果</param>
        /// <param name="sourceRect">截图区域位置</param>
        /// <param name="processingTimeSeconds">处理时间（秒）</param>
        public static void ShowOCRResult(string ocrResult, Rectangle sourceRect, double processingTimeSeconds)
        {
            try
            {
                var resultWindow = new OCRResultWindow(ocrResult, sourceRect, processingTimeSeconds);
                resultWindow.Show();
                resultWindow.BringToFront();
                resultWindow.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"显示结果窗口失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
        
        /// <summary>
        /// 显示错误信息悬浮窗
        /// </summary>
        /// <param name="errorMessage">错误信息</param>
        /// <param name="sourceRect">截图区域位置</param>
        public static void ShowError(string errorMessage, Rectangle sourceRect)
        {
            string fullErrorMessage = $"OCR识别失败\n\n{errorMessage}\n\n请确保:\n1. Ollama服务正在运行 (ollama serve)\n2. 模型已安装 (ollama pull benhaotang/Nanonets-OCR-s:latest)\n3. 服务地址正确 (http://localhost:11434)";
            ShowOCRResult(fullErrorMessage, sourceRect);
        }
    }
}