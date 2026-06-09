# 京韵智析：Streamlit Community Cloud 部署包

## 上传前

把前四问的 Excel/CSV 结果复制到 `data` 文件夹。

## GitHub 仓库结构

```text
streamlit_app.py
requirements.txt
.streamlit/config.toml
data/
  第一问结果.xlsx
  第二问网络指标.xlsx
  第二问关系边表.xlsx
  第三问主题结果.xlsx
  第四问剧本叙事指标.xlsx
  第四问场次叙事指标.xlsx
```

## Streamlit Community Cloud 设置

- Repository：选择上传本项目的 GitHub 仓库
- Branch：main
- Main file path：streamlit_app.py
- Python version：3.11

部署后会得到 `https://你的应用名.streamlit.app`。
