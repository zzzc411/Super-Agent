# 空白 Word 文档

生成一个空白的 `.docx` 文件：一页 A4、无正文文字、无页眉页脚、无表格图片。

## 直接使用

仓库内已附生成好的文件：[`blank.docx`](blank.docx)。下载后即可用 Word / WPS 打开，得到一页空白文档，可直接开始输入。

## 重新生成

```bash
pip install python-docx
python3 make_blank_docx.py            # 写出同目录的 blank.docx
python3 make_blank_docx.py out.docx   # 或指定输出路径
```

文档页边距取 python-docx 默认模板的值（上下 2.54cm、左右 3.18cm），与 Word 新建文档一致。

## 校验

```bash
pip install python-docx
python3 verify_blank_docx.py blank.docx
```

校验项：正文没有文字、没有表格和图片、只有一个节、页面为 A4 纵向、页眉页脚为空。全部通过时以退出码 0 结束，否则打印具体问题并返回 1。
