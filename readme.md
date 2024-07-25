<!-- @format -->

# PDF2DOCX Pro

[English](#english) | [中文](#中文)

![python](https://img.shields.io/badge/python-≥3.6-blue)
![license](https://img.shields.io/badge/license-MIT-green)

Advanced Python toolkit for lightning-fast PDF to DOCX conversion. Seamlessly handles text and image-based PDFs with intelligent processing and OCR integration.

## English

### Features

-   Fast conversion of text-based PDFs using PyMuPDF
-   OCR processing for image-based PDFs using PaddleOCR
-   Multi-threaded processing for improved efficiency
-   Automatic detection and fixing of problematic conversions
-   Supports both English and Chinese documents
-   Intelligent content validation and reprocessing

### Advantages

-   High efficiency: Uses PyMuPDF for quick conversion of text PDFs
-   Versatility: Employs OCR for image-based PDFs, ensuring comprehensive coverage
-   Intelligent processing: Automatically detects and reprocesses problematic conversions
-   User-friendly: Simple command-line interface for easy operation
-   Robust: Handles various PDF types and formats

### Installation

1. Clone the repository:

git clone https://github.com/yourusername/PDF2DOCX-Pro.git
cd PDF2DOCX-Pro

2. Create a virtual environment (optional but recommended):

python -m venv venv
source venv/bin/activate # On Windows use venv\Scripts\activate

3. Install the required packages:

pip install -r requirements.txt

### Usage

Run the script with the following command:

python pdf2docx_pro.py <input_folder> <output_folder> [--max_workers MAX_WORKERS] [--size_threshold SIZE_THRESHOLD]

Arguments:

-   `input_folder`: Path to the folder containing PDF files
-   `output_folder`: Path to save the converted DOCX files
-   `--max_workers`: (Optional) Maximum number of worker threads (default: 4)
-   `--size_threshold`: (Optional) Small file size threshold in KB (default: 15)

Example:

python pdf2docx_pro.py ./pdf_files ./docx_files --max_workers 6 --size_threshold 20

### Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 中文

### 特性

-   使用 PyMuPDF 快速转换文本类 PDF
-   使用 PaddleOCR 对图片类 PDF 进行 OCR 处理
-   多线程处理提高效率
-   自动检测并修复问题转换
-   支持英文和中文文档
-   智能内容验证和重新处理

### 优势

-   高效率：使用 PyMuPDF 快速转换文本 PDF
-   通用性：对图片类 PDF 采用 OCR 处理，确保全面覆盖
-   智能处理：自动检测并重新处理问题转换
-   用户友好：简单的命令行界面，操作便捷
-   健壮性：能够处理各种类型和格式的 PDF

### 安装

1. 克隆仓库：

git clone https://github.com/yourusername/PDF2DOCX-Pro.git
cd PDF2DOCX-Pro

2. 创建虚拟环境（可选但推荐）：

python -m venv venv
source venv/bin/activate # Windows 下使用 venv\Scripts\activate

3. 安装所需包：

pip install -r requirements.txt

### 使用方法

使用以下命令运行脚本：

python pdf2docx_pro.py <input_folder> <output_folder> [--max_workers MAX_WORKERS] [--size_threshold SIZE_THRESHOLD]

参数：

-   `input_folder`: 包含 PDF 文件的文件夹路径
-   `output_folder`: 保存转换后 DOCX 文件的文件夹路径
-   `--max_workers`: （可选）最大工作线程数（默认：4）
-   `--size_threshold`: （可选）小文件大小阈值，单位为 KB（默认：15）

示例：

python pdf2docx_pro.py ./pdf_files ./docx_files --max_workers 6 --size_threshold 20

### 贡献

欢迎贡献！请随时提交 Pull Request。

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 将您的更改推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

### 许可证

本项目采用 MIT 许可证 - 详情请见 [LICENSE](LICENSE) 文件。
