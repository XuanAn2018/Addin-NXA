[Hướng dẫn sử dụng chi tiết Add-in NXA](https://nnct2023.github.io/Addin-NXA/): Trợ lý AI Đa Năng Nâng Tầm Trải Nghiệm Excel của Bạn


# 🚀 Addin-NXA: Tiện Ích AI Đa Năng Cho Excel

Addin-NXA là tiện ích mã nguồn mở từ cộng đồng Việt, tích hợp các nền tảng AI hàng đầu (Gemini, Cloudflare AI, OpenRouter) vào Excel, mang lại sự linh hoạt và các hàm AI mạnh mẽ, tối ưu cho người dùng Việt.

## 🌟 Các Điểm Nổi Bật

| 🏷️ Tính năng                     | 📝 Mô tả                                                                 | 💡 Chi tiết nổi bật                                                                 |
|----------------------------------|-------------------------------------------------------------------------|------------------------------------------------------------------------------------|
| 🌐 Tích hợp đa nền tảng          | Kết nối Gemini, Cloudflare AI, OpenRouter.                              | Hỗ trợ Claude, Llama, Mistral, GPT-OSS (@cf/openai/gpt-oss-120b, @cf/openai/gpt-oss-20b). |
| 📚 Kho hàm AI toàn diện          | Bộ hàm hỏi đáp, phân tích, dịch thuật.                                  | NXA_AskGemini, NXA_AIMemories, NXA_Insights, NXA_AITranslator, NXA_Explain.        |
| 🇻🇳 Tối ưu cho người Việt        | Hàm tiện ích địa phương.                                               | Tạo QR ngân hàng (NNCT_QR_Bank), phân tích ngành may (Lean Manufacturing).         |
| 💬 Giao diện trực quan           | Chat AI trong Excel.                                                   | Trải nghiệm như add-in quốc tế (ChatGPT for Excel).                                |
| ⚙️ Linh hoạt cấu hình            | Tùy chỉnh API key/model.                                               | Lưu cấu hình vào config.txt qua NXA_ZConfigAPI.                                    |

## 📋 Yêu Cầu

| 🏷️ Yêu cầu                       | 📝 Mô tả                                                                 | 💡 Chi tiết                                                                 |
|----------------------------------|-------------------------------------------------------------------------|-----------------------------------------------------------------------------|
| 📊 Excel                         | Phiên bản hỗ trợ VBA (2010+).                                          | Kích hoạt Developer tab để dùng VBA.                                        |
| 🌐 Internet                      | Kết nối ổn định để gọi API.                                            | Đảm bảo không bị chặn bởi firewall.                                         |
| 🔑 Tài khoản Cloudflare/Gemini   | Tài khoản miễn phí để lấy API key.                                     | Đăng ký tại [Google AI Studio](https://makersuite.google.com/app/apikey) hoặc Cloudflare. |

## 🔑 Chuẩn Bị API Key

| 🏷️ Bước                          | 📝 Mô tả                                                                 | 💡 Hành động cụ thể                                                         |
|----------------------------------|-------------------------------------------------------------------------|-----------------------------------------------------------------------------|
| 🌐 Lấy Account ID                | Sao chép từ Cloudflare Dashboard.                                      | Đăng nhập [Cloudflare Dashboard](https://dash.cloudflare.com) > Overview.   |
| 🔐 Tạo API Token                 | Lấy Global API Key từ Cloudflare.                                      | Profile > API Tokens > Global API Key > View/Create.                        |
| 📋 Lấy Gemini API Key            | Đăng ký và lấy từ Google AI Studio.                                    | Truy cập [Google AI Studio](https://makersuite.google.com/app/apikey).      |

## 🛠️ Triển Khai Trong Excel

| 🏷️ Bước                          | 📝 Mô tả                                                                 | 💡 Hành động cụ thể                                                         |
|----------------------------------|-------------------------------------------------------------------------|-----------------------------------------------------------------------------|
| 📥 Tải Add-in                   | Tải file NXA.xlam từ GitHub.                                           | Truy cập [Releases](https://github.com/XuanAn2018/Addin-NXA/releases).      |
| ➕ Kích hoạt Add-in              | Thêm NXA.xlam vào Excel.                                               | File > Options > Add-ins > Excel Add-ins > Browse > Chọn NXA.xlam > OK.     |
| 🧩 Cài JsonConverter             | Thêm thư viện VBA-JSON.                                                | Alt+F11 > Import [JsonConverter.bas](https://github.com/VBA-tools/VBA-JSON).|
| 📚 Thêm References               | Kích hoạt thư viện VBA.                                                | Tools > References > Tích Microsoft XML v6.0, Microsoft Scripting Runtime.   |
| 💾 Lưu File                      | Lưu dưới định dạng *.xlsm.                                             | File > Save As > Excel Macro-Enabled Workbook (*.xlsm).                     |

## ⚙️ Cấu Hình Lần Đầu

| 🏷️ Bước                          | 📝 Mô tả                                                                 | 💡 Hành động cụ thể                                                         |
|----------------------------------|-------------------------------------------------------------------------|-----------------------------------------------------------------------------|
| 🚀 Chạy Test                    | Chạy Sub TestCloudflareAIAddin.                                        | Alt+F11 > Chọn Sub TestCloudflareAIAddin > F5.                              |
| 🔑 Nhập Account ID              | Dán Account ID từ Cloudflare.                                          | Nhập vào hộp thoại hiện ra.                                                |
| 🤖 Chọn Model                   | Chọn hoặc nhập Model ID.                                               | Nhập số (1, 2,...) hoặc Model ID (ví dụ: @cf/openai/gpt-oss-120b).          |
| 🔐 Nhập API Token               | Dán Global API Key hoặc Gemini Key.                                    | Nhập vào hộp thoại, lưu vào Documents\ChatLogs\CloudflareAI\config.txt.     |

## 🚀 Sử Dụng Hàm AI

| 🏷️ Chức năng                     | 📝 Mô tả                                                                 | 💡 Cú pháp hàm                                                      | 📊 Kết quả                              |
|----------------------------------|-------------------------------------------------------------------------|--------------------------------------------------------------------|----------------------------------------|
| ❓ Hỏi đáp                       | Gửi câu hỏi đến Gemini.                                                | =NXA_AskGemini("Tổng doanh thu 2025?")                             | Ví dụ: 500 triệu VND                   |
| 🌐 Dịch thuật                    | Dịch văn bản sang ngôn ngữ mong muốn.                                  | =NCC_AITranslator(A2, "vi")                                       | Ví dụ: Xin chào, bạn khỏe không?       |
| 📊 Phân tích dữ liệu             | Sinh insight từ dữ liệu.                                               | =NXA_Insights(A1:B10, "Phân tích xu hướng")                       | Xu hướng tăng 15% quý 4                |
| 💬 Trò chuyện ngữ cảnh           | Duy trì lịch sử hội thoại.                                             | =CF_AIMemories("Xin chào, bạn nhớ tôi?", TRUE)                   | AI trả lời dựa trên ngữ cảnh trước      |

**Ví dụ**: Ô A1: "What is the capital of Vietnam?", ô B1: `=NXA_AskGemini(A1)` → Kết quả: "Hà Nội".

## ⚠️ Lưu Ý Quan Trọng

| 🏷️ Lưu ý                        | 📝 Mô tả                                                                 | 💡 Chi tiết                                                                 |
|----------------------------------|-------------------------------------------------------------------------|-----------------------------------------------------------------------------|
| 🛡️ Bảo mật API Key              | Không chia sẻ key công khai.                                           | Lưu key an toàn, chỉ dùng trong môi trường đáng tin cậy.                    |
| 📡 Kết nối Internet              | Cần mạng ổn định để gọi API.                                           | Kiểm tra firewall/proxy nếu gặp lỗi kết nối.                               |
| 🧪 Thử nghiệm an toàn            | Dùng dữ liệu không nhạy cảm.                                           | Tránh dữ liệu khách hàng trước khi kiểm tra đầy đủ.                         |
| 📊 Định dạng dữ liệu             | Đảm bảo dữ liệu đầu vào đúng.                                          | Chuỗi văn bản hoặc vùng có header rõ ràng (ví dụ: A1:B10).                  |
| ⏱️ Thời gian phản hồi            | Có thể mất 2-5 giây tùy model.                                         | Chọn model nhẹ (như gemini-1.5-flash) nếu cần tốc độ.                       |

## 📚 Tài Nguyên Xác Minh
- [GitHub Addin-NXA](https://github.com/XuanAn2018/Addin-NXA)
- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)
- [Google AI Studio (Gemini API)](https://makersuite.google.com/app/apikey)
- [Cloudflare Workers AI Docs](https://developers.cloudflare.com/workers-ai/)
- [Config API Guide](https://nnct2023.github.io/Addin-NXA/docs/setup/ConfigAPI.html)

