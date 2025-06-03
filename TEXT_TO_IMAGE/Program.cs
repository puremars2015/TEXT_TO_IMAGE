using SkiaSharp;
using System;
using System.IO;
using System.Net.Sockets;
using System.Text;

namespace TEXT_TO_IMAGE;
class Program
{
    private const string DEFAULT_FONT_PATH = "/System/Library/Fonts/STHeiti Medium.ttc";
    private const string ZEBRA_IP = "192.168.1.45";
    private const int ZEBRA_PORT = 9100;

    static void Main()
    {
        string text = "標籤列印測試\n（中文OK）";
        string outputPath = "label_text_mac.png";

        try
        {
            // // 1. 文字轉圖片
            // ConvertTextToImage(text, outputPath, 48);
            // Console.WriteLine("✅ 中文圖片儲存完成：" + outputPath);



            outputPath = "/Users/maenqi/Projects/TEXT_TO_IMAGE/TEXT_TO_IMAGE/bin/Debug/net6.0/label_capture_bw.png";

            // 2. 讀取圖片並準備打印
            using var scaledBitmap = LoadAndScaleImage(outputPath, 2.0);
            
            // 裁剪圖片到固定大小 (例如 600x400 像素)，居中裁剪
            int targetWidth = 300;
            int targetHeight = 200;
            using var bitmap = CropImageToSize(scaledBitmap, targetWidth, targetHeight, true);
            
            // 3. 轉換為二進位碼
            string binaryData = ConvertImageToZplHex(bitmap);
            
            // 4. 建立ZPL指令
            string zplCommand = GenerateZplCommand(bitmap.Width, bitmap.Height, binaryData);
            
            // 5. 傳送到印表機
            SendZplToPrinter(zplCommand, ZEBRA_IP, ZEBRA_PORT);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 處理過程發生錯誤: {ex.Message}");
        }
    }

    /// <summary>
    /// 將文字轉換為圖片並儲存 (支援多行文字)
    /// </summary>
    private static void ConvertTextToImage(string text, string outputPath, int fontSize)
    {
        if (!File.Exists(DEFAULT_FONT_PATH))
        {
            throw new FileNotFoundException($"找不到字型檔案：{DEFAULT_FONT_PATH}");
        }

        var typeface = SKTypeface.FromFile(DEFAULT_FONT_PATH);
        var paint = new SKPaint
        {
            Typeface = typeface,
            TextSize = fontSize,
            IsAntialias = true,
            Color = SKColors.Black
        };
        
        // 分割文字為多行
        string[] lines = text.Split('\n');
        
        // 計算每行文字的寬度和總高度
        float maxWidth = 0;
        float totalHeight = 0;
        var lineHeights = new float[lines.Length];
        
        for (int i = 0; i < lines.Length; i++)
        {
            var bounds = new SKRect();
            float lineWidth = paint.MeasureText(lines[i], ref bounds);
            maxWidth = Math.Max(maxWidth, lineWidth);
            lineHeights[i] = bounds.Height;
            totalHeight += bounds.Height;
        }
        
        // 行間距 (20% 的字體大小)
        float lineSpacing = fontSize * 0.2f;
        totalHeight += lineSpacing * (lines.Length - 1);
        
        // 增加邊距 (每邊增加原尺寸的50%，長寬總共增加兩倍)
        float margin = fontSize * 2;  // 使用字體大小的2倍作為邊距
        float totalWidth = maxWidth + margin * 2;
        totalHeight += margin * 2;
        
        // 創建位圖並繪製文字
        using var bitmap = new SKBitmap((int)Math.Ceiling(totalWidth), (int)Math.Ceiling(totalHeight));
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);
        
        // 繪製每行文字，起始位置增加邊距
        float y = margin;
        for (int i = 0; i < lines.Length; i++)
        {
            var bounds = new SKRect();
            paint.MeasureText(lines[i], ref bounds);
            
            // 調整 y 座標使文字垂直置中於其邊界
            y -= bounds.Top;
            
            canvas.DrawText(lines[i], margin, y, paint);
            
            // 設定下一行的 y 座標
            y += bounds.Height + lineSpacing;
        }
        
        canvas.Flush();

        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        using var stream = File.OpenWrite(outputPath);
        data.SaveTo(stream);
    }

    /// <summary>
    /// 載入圖片並進行縮放
    /// </summary>
    private static SKBitmap LoadAndScaleImage(string imagePath, double scale)
    {
        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"找不到圖片檔案：{imagePath}");
        }

        Console.WriteLine($"讀取圖片: {imagePath}");
        using var originalBitmap = SKBitmap.Decode(imagePath);
        if (originalBitmap == null)
        {
            throw new InvalidOperationException($"無法解碼圖片: {imagePath}");
        }
        
        Console.WriteLine($"原始圖片大小: {originalBitmap.Width}x{originalBitmap.Height}");
        int scaledWidth = (int)(originalBitmap.Width * scale);
        int scaledHeight = (int)(originalBitmap.Height * scale);
        Console.WriteLine($"縮放後大小: {scaledWidth}x{scaledHeight}");
        
        var scaledBitmap = new SKBitmap(scaledWidth, scaledHeight);
        using (var canvas = new SKCanvas(scaledBitmap))
        {
            canvas.Clear(SKColors.White);
            var destRect = new SKRect(0, 0, scaledWidth, scaledHeight);
            canvas.DrawBitmap(originalBitmap, destRect);
        }
        
        return scaledBitmap;
    }

    /// <summary>
    /// 將圖片轉換為 ZPL 支援的十六進制字串格式
    /// </summary>
    private static string ConvertImageToZplHex(SKBitmap bitmap)
    {
        Console.WriteLine($"開始將圖片轉換為ZPL二進制碼，圖片尺寸: {bitmap.Width}x{bitmap.Height}");
        
        // 轉換為黑白圖像
        using var blackWhiteBitmap = new SKBitmap(bitmap.Width, bitmap.Height, SKColorType.Gray8, SKAlphaType.Opaque);
        using (var canvas = new SKCanvas(blackWhiteBitmap))
        {
            canvas.DrawBitmap(bitmap, 0, 0);
        }
        
        // 檢查圖片大小，過大的圖片可能無法正常處理
        if (bitmap.Width > 1500 || bitmap.Height > 1500)
        {
            Console.WriteLine($"⚠ 警告: 圖片尺寸過大 ({bitmap.Width}x{bitmap.Height})，可能導致處理問題或打印失敗");
        }
        
        int bytesPerRow = (bitmap.Width + 7) / 8;
        var hexString = new StringBuilder();
        int blackPixelCount = 0;
        int totalPixels = bitmap.Width * bitmap.Height;
        
        for (int y = 0; y < bitmap.Height; y++)
        {
            byte[] rowBytes = new byte[bytesPerRow];
            
            for (int x = 0; x < bitmap.Width; x++)
            {
                SKColor pixel = blackWhiteBitmap.GetPixel(x, y);
                byte grayValue = (byte)((pixel.Red * 0.3) + (pixel.Green * 0.59) + (pixel.Blue * 0.11));
                
                if (grayValue < 128)
                {
                    int byteIndex = x / 8;
                    int bitIndex = 7 - (x % 8);
                    rowBytes[byteIndex] |= (byte)(1 << bitIndex);
                    blackPixelCount++;
                }
            }
            
            foreach (byte b in rowBytes)
            {
                hexString.Append(b.ToString("X2"));
            }
        }
        
        double blackPixelPercentage = (double)blackPixelCount / totalPixels * 100;
        Console.WriteLine($"圖片分析: 黑色像素佔比 {blackPixelPercentage:F2}%");
        
        // 如果黑色像素比例過低，可能無法看到印刷效果
        if (blackPixelPercentage < 1)
        {
            Console.WriteLine($"⚠ 警告: 黑色像素比例過低，打印結果可能不明顯");
        }
        
        string result = hexString.ToString();
        Console.WriteLine($"ZPL二進制碼生成完成，長度: {result.Length} 字符");
        
        return result;
    }

    /// <summary>
    /// 生成完整的 ZPL 指令
    /// </summary>
    private static string GenerateZplCommand(int imageWidth, int imageHeight, string hexData)
    {
        int byteCount = hexData.Length / 2;
        int bytesPerRow = (imageWidth + 7) / 8;
        
        Console.WriteLine($"生成ZPL指令: 寬度={imageWidth}, 高度={imageHeight}, 數據長度={byteCount}字節");
        
        // 調整為更接近邊緣的位置，確保圖片可見
        string zplCommand = "^XA" +
               $"^FO50,50" +  // 將位置調整為左上角更靠近邊緣的位置
               $"^GFA,{byteCount},{byteCount},{bytesPerRow},{hexData}" +
               "^FS" +
               "^XZ";
               
        Console.WriteLine($"ZPL指令總長度: {zplCommand.Length} 字符");
        return zplCommand;
    }

    /// <summary>
    /// 發送 ZPL 指令到印表機
    /// </summary>
    private static void SendZplToPrinter(string zplCommand, string ipAddress, int port)
    {
        try
        {
            using var client = new TcpClient();
            client.Connect(ipAddress, port);
            using var stream = client.GetStream();
            byte[] zplBytes = Encoding.UTF8.GetBytes(zplCommand);
            stream.Write(zplBytes, 0, zplBytes.Length);
            Console.WriteLine($"✅ 已傳送列印指令至 Zebra 標籤機 (IP: {ipAddress})");
        }
        catch (Exception ex)
        {
            throw new Exception($"傳送列印指令失敗: {ex.Message}");
        }
    }

    /// <summary>
    /// 裁剪圖片至指定大小，可選擇居中裁剪或從特定位置裁剪
    /// </summary>
    /// <param name="bitmap">原始圖片</param>
    /// <param name="targetWidth">目標寬度</param>
    /// <param name="targetHeight">目標高度</param>
    /// <param name="centerCrop">是否居中裁剪，若為 false 則從左上角開始裁剪</param>
    /// <returns>裁剪後的圖片</returns>
    private static SKBitmap CropImageToSize(SKBitmap bitmap, int targetWidth, int targetHeight, bool centerCrop = true)
    {
        Console.WriteLine($"裁剪圖片: 原始大小 {bitmap.Width}x{bitmap.Height} => 目標大小 {targetWidth}x{targetHeight}");
        
        // 如果原始圖片小於目標尺寸，則不裁剪，而是創建一個白色背景並將圖片置中
        if (bitmap.Width <= targetWidth && bitmap.Height <= targetHeight)
        {
            Console.WriteLine("原始圖片小於目標尺寸，將在白色背景上置中");
            var resultBitmap = new SKBitmap(targetWidth, targetHeight);
            using var canvas = new SKCanvas(resultBitmap);
            canvas.Clear(SKColors.White);
            
            // 計算居中位置
            int centerX = (targetWidth - bitmap.Width) / 2;
            int centerY = (targetHeight - bitmap.Height) / 2;
            
            canvas.DrawBitmap(bitmap, centerX, centerY);
            return resultBitmap;
        }
        
        // 計算裁剪區域
        int x = 0, y = 0;
        if (centerCrop)
        {
            // 居中裁剪
            x = Math.Max(0, (bitmap.Width - targetWidth) / 2);
            y = Math.Max(0, (bitmap.Height - targetHeight) / 2);
        }
        
        // 確保裁剪區域不會超出原始圖片範圍
        int realWidth = Math.Min(targetWidth, bitmap.Width - x);
        int realHeight = Math.Min(targetHeight, bitmap.Height - y);
        
        // 創建裁剪區域
        var sourceRect = new SKRectI(x, y, x + realWidth, y + realHeight);
        
        // 提取裁剪區域並創建新位圖
        var croppedBitmap = new SKBitmap(targetWidth, targetHeight);
        using (var canvas = new SKCanvas(croppedBitmap))
        {
            canvas.Clear(SKColors.White);
            canvas.DrawBitmap(
                bitmap, 
                sourceRect, 
                new SKRect(0, 0, realWidth, realHeight)
            );
        }
        
        Console.WriteLine($"裁剪完成: 最終大小 {croppedBitmap.Width}x{croppedBitmap.Height}");
        return croppedBitmap;
    }
}

