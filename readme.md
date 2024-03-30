## 使用 Python 搜索 Excel 文件中的关键词

这个 Python 程序可以帮助您在 Excel 文件中搜索指定的关键词。

### 使用方法

1. **安装依赖库**

   ```
   pip install openpyxl
   ```

2. **修改代码**

   * 修改 `keywords_to_search` 变量，设置要搜索的关键词列表。
   * 修改 `folder_to_search` 变量，设置包含 Excel 文件的文件夹路径。

3. **运行程序**

   ```
   python search_keywords_in_excel.py
   ```

### 示例

假设我们要搜索包含 "Python" 和 "Excel" 关键词的 Excel 文件，并将结果输出到控制台。

**1. 修改代码**

```python
keywords_to_search = ["Python", "Excel"]
folder_to_search = "path/to/folder"
```

**2. 运行程序**

```
python search_keywords_in_excel.py
```

**3. 输出结果**

```
关键词 'Python', 'Excel' 在文件 'file_name.xlsx'的数据为: ['This is a cell with Python and Excel', 'Another cell with Python']
```

### 其他说明

* 此程序可以递归搜索文件夹中的所有 Excel 文件。
* 程序会输出包含关键词的单元格所在的行数据。
* 您可以根据需要修改程序以满足您的特定需求。

### 改进建议

* 添加对多个 Excel 工作表的支持。
* 添加对不同格式 Excel 文件的支持。
* 添加对模糊搜索的支持。
* 添加对正则表达式搜索的支持。
