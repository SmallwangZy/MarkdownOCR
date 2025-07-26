using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace MarkdownOCR
{
    /// <summary>
    /// Ollama API响应模型
    /// </summary>
    public class OllamaResponse
    {
        public string model { get; set; } = string.Empty;
        public string created_at { get; set; } = string.Empty;
        public string response { get; set; } = string.Empty;
        public bool done { get; set; }
        public int[] context { get; set; } = Array.Empty<int>();
        public long total_duration { get; set; }
        public long load_duration { get; set; }
        public int prompt_eval_count { get; set; }
        public long prompt_eval_duration { get; set; }
        public int eval_count { get; set; }
        public long eval_duration { get; set; }
    }

    /// <summary>
    /// Ollama API请求模型
    /// </summary>
    public class OllamaRequest
    {
        public string model { get; set; } = string.Empty;
        public string prompt { get; set; } = string.Empty;
        public bool stream { get; set; } = false;
        public string[] images { get; set; } = Array.Empty<string>();
    }

    /// <summary>
    /// Ollama客户端，负责与本地Ollama后端通信
    /// </summary>
    public static class OllamaClient
    {
        // Ollama配置常量
        private const string OLLAMA_BASE_URL = "http://localhost:11434";
        private const string OLLAMA_API_ENDPOINT = "/api/generate";
        private const string OLLAMA_MODEL = "benhaotang/Nanonets-OCR-s:latest";
        private const string OCR_PROMPT = "请识别图片中的文本，以Markdown格式返回识别出的文本内容，不要添加任何解释或格式化。";
        private const int REQUEST_TIMEOUT_SECONDS = 60;

        private static readonly HttpClient _httpClient = new HttpClient()
        {
            Timeout = TimeSpan.FromSeconds(REQUEST_TIMEOUT_SECONDS)
        };

        /// <summary>
        /// 向Ollama发送OCR识别请求
        /// </summary>
        /// <param name="base64Image">Base64编码的图像</param>
        /// <returns>包含识别结果和处理时间的元组</returns>
        public static async Task<(string result, double processingTimeSeconds)> RecognizeTextAsync(string base64Image)
        {
            if (string.IsNullOrEmpty(base64Image))
                throw new ArgumentException("Base64图像数据不能为空", nameof(base64Image));

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                // 构建请求数据
                var request = new OllamaRequest
                {
                    model = OLLAMA_MODEL,
                    prompt = OCR_PROMPT,
                    stream = false,
                    images = new[] { base64Image }
                };

                // 序列化请求为JSON
                string jsonRequest = JsonSerializer.Serialize(request, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                // 创建HTTP请求内容
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                // 发送POST请求到Ollama API
                string apiUrl = OLLAMA_BASE_URL + OLLAMA_API_ENDPOINT;
                Console.WriteLine($"正在向Ollama发送OCR请求: {apiUrl}");
                Console.WriteLine($"使用模型: {OLLAMA_MODEL}");

                HttpResponseMessage response = await _httpClient.PostAsync(apiUrl, content);

                // 检查响应状态
                if (!response.IsSuccessStatusCode)
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"Ollama API请求失败 (状态码: {response.StatusCode}): {errorContent}");
                }

                // 读取响应内容
                string responseContent = await response.Content.ReadAsStringAsync();
                
                // 解析JSON响应
                var ollamaResponse = JsonSerializer.Deserialize<OllamaResponse>(responseContent, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                if (ollamaResponse == null)
                {
                    throw new InvalidOperationException("无法解析Ollama API响应");
                }

                // 检查响应是否完成
                if (!ollamaResponse.done)
                {
                    throw new InvalidOperationException("Ollama API响应未完成");
                }

                stopwatch.Stop();
                double processingTime = stopwatch.ElapsedMilliseconds / 1000.0;
                
                Console.WriteLine($"OCR处理完成，耗时: {processingTime:F2} 秒");
                
                // 返回识别结果和处理时间
                return (ollamaResponse.response?.Trim() ?? string.Empty, processingTime);
            }
            catch (HttpRequestException ex)
            {
                stopwatch.Stop();
                throw new InvalidOperationException($"无法连接到Ollama后端服务 ({OLLAMA_BASE_URL}): {ex.Message}\n\n请确保Ollama服务正在运行，并且模型 '{OLLAMA_MODEL}' 已安装。", ex);
            }
            catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException)
            {
                stopwatch.Stop();
                throw new TimeoutException($"Ollama API请求超时 ({REQUEST_TIMEOUT_SECONDS}秒)。请检查网络连接和服务状态。", ex);
            }
            catch (JsonException ex)
            {
                stopwatch.Stop();
                throw new InvalidOperationException($"解析Ollama API响应失败: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                throw new InvalidOperationException($"OCR识别过程中发生未知错误: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 检查Ollama服务是否可用
        /// </summary>
        /// <returns>服务是否可用</returns>
        public static async Task<bool> IsServiceAvailableAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync(OLLAMA_BASE_URL);
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取Ollama服务信息
        /// </summary>
        /// <returns>服务信息字符串</returns>
        public static string GetServiceInfo()
        {
            return $"Ollama服务地址: {OLLAMA_BASE_URL}\n" +
                   $"API端点: {OLLAMA_API_ENDPOINT}\n" +
                   $"使用模型: {OLLAMA_MODEL}\n" +
                   $"请求超时: {REQUEST_TIMEOUT_SECONDS}秒";
        }

        /// <summary>
        /// 释放HTTP客户端资源
        /// </summary>
        public static void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}