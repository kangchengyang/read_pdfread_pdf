git clone git@github.com:kangchengyang/read_pdfread_pdf.git

uv sync

## 环境导出本地
# 将 requirements.txt 中的包下载到 packages 文件夹
uv pip download -r requirements.txt -d ./packages

# 离线安装
uv pip install -r requirements.txt --find-links ./packages
