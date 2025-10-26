# datafeed.py 新增功能文档

## 概述

本次更新为 `datafeed.py` 添加了以下主要功能模块：

1. **Excel处理模块** (`ExcelProcessor`)
2. **数据验证工具** (`DataValidator`)
3. **数据库-Excel交互功能** (在 `Datafeed` 类中)
4. **重构的查询方法** (改进的 `query_nearest_after` 和 `query_nearest_before`)

---

## 1. Excel处理模块 (ExcelProcessor)

### 功能描述

`ExcelProcessor` 类提供了完整的Excel/CSV文件处理功能，包括：
- 读取CSV和XLSX文件
- 转换cross-section数据为数据库三列表格式
- 批量文件操作
- 文件夹分类和目录树组织
- 自动生成日志和警告

### 主要方法

#### `__init__(log_dir=None)`
初始化Excel处理器
- **参数**: 
  - `log_dir`: 日志文件保存目录

#### `read_file(filepath, **kwargs)`
读取CSV或XLSX文件并转换为DataFrame
- **参数**:
  - `filepath`: 文件路径
  - `**kwargs`: 传递给pandas读取函数的额外参数
- **返回**: DataFrame或None

#### `save_dataframe(df, filepath, file_format='csv', **kwargs)`
保存DataFrame到文件
- **参数**:
  - `df`: 要保存的DataFrame
  - `filepath`: 保存路径
  - `file_format`: 文件格式 ('csv' 或 'xlsx')
- **返回**: 是否成功保存 (bool)

#### `transform_to_db_format(df, key_columns, value_column, metric_name=None)`
将cross-section数据转化为数据库三列表格式
- **参数**:
  - `df`: 输入DataFrame
  - `key_columns`: 作为key的列名列表
  - `value_column`: 包含值的列名
  - `metric_name`: metric名称
- **返回**: 转换后的DataFrame

#### `batch_read_files(directory, pattern='*', recursive=False, **kwargs)`
批量读取目录中的文件
- **参数**:
  - `directory`: 目录路径
  - `pattern`: 文件匹配模式
  - `recursive`: 是否递归读取子目录
- **返回**: 字典，键为文件路径，值为DataFrame

#### `organize_output(dataframes, output_dir, categorize_by=None)`
批量保存文件并按分类组织
- **参数**:
  - `dataframes`: 字典，键为原始文件名，值为DataFrame
  - `output_dir`: 输出目录
  - `categorize_by`: 分类函数
- **返回**: 目录树结构字典

### 使用示例

```python
from betalens.datafeed import ExcelProcessor

# 初始化处理器
processor = ExcelProcessor(log_dir='./logs')

# 读取文件
df = processor.read_file('data.csv')

# 转换为数据库格式
db_format = processor.transform_to_db_format(
    df,
    key_columns=['code', 'name', 'date'],
    value_column='close_price',
    metric_name='收盘价'
)

# 保存文件
processor.save_dataframe(db_format, 'output.csv')

# 批量读取文件
files = processor.batch_read_files('./data', pattern='*.csv', recursive=True)
```

---

## 2. 数据验证工具 (DataValidator)

### 功能描述

`DataValidator` 类提供了全面的数据质量检查功能：
- 检查空值、NaN、None值
- 验证日期列（重复、排序、频率）
- 检查数据类型
- 多种缺失值处理方法
- 综合数据检查

### 主要方法

#### `__init__(logger=None)`
初始化数据验证器
- **参数**: 
  - `logger`: 日志记录器（可选）

#### `check_null_values(df, columns=None)`
检查DataFrame中的空值、NaN、None值
- **参数**:
  - `df`: 待检查的DataFrame
  - `columns`: 要检查的列名列表
- **返回**: 字典，键为列名，值为空值数量

#### `check_date_column(df, date_column, check_duplicates=True, check_sorting=True, check_frequency=False, expected_freq=None)`
检查日期列的各种问题
- **参数**:
  - `df`: 待检查的DataFrame
  - `date_column`: 日期列名
  - `check_duplicates`: 是否检查重复日期
  - `check_sorting`: 是否检查排序
  - `check_frequency`: 是否检查频率
  - `expected_freq`: 期望的频率（如 'D', 'W', 'M'）
- **返回**: 包含检查结果的字典

#### `check_data_types(df, expected_types)`
检查列的数据类型是否符合预期
- **参数**:
  - `df`: 待检查的DataFrame
  - `expected_types`: 字典，键为列名，值为期望的类型
- **返回**: 字典，键为列名，值为是否匹配

#### `handle_missing_values(df, method='drop', columns=None, fill_value=None, **kwargs)`
处理缺失值
- **参数**:
  - `df`: 待处理的DataFrame
  - `method`: 处理方法
    - `'drop'`: 删除包含空值的行
    - `'ffill'`: 前向填充
    - `'bfill'`: 后向填充
    - `'fill'`: 使用指定值填充
    - `'raise'`: 抛出错误
  - `columns`: 要处理的列
  - `fill_value`: 填充值（method='fill'时使用）
- **返回**: 处理后的DataFrame

#### `comprehensive_check(df, date_columns=None, required_columns=None)`
综合检查DataFrame
- **参数**:
  - `df`: 待检查的DataFrame
  - `date_columns`: 日期列列表
  - `required_columns`: 必须存在的列
- **返回**: 包含所有检查结果的字典

### 使用示例

```python
from betalens.datafeed import DataValidator

# 初始化验证器
validator = DataValidator()

# 检查空值
null_counts = validator.check_null_values(df)

# 检查日期列
date_check = validator.check_date_column(
    df, 
    'datetime',
    check_duplicates=True,
    check_sorting=True
)

# 处理缺失值
clean_df = validator.handle_missing_values(
    df, 
    method='ffill',
    columns=['value']
)

# 综合检查
results = validator.comprehensive_check(
    df,
    date_columns=['datetime'],
    required_columns=['code', 'datetime', 'value', 'metric']
)
```

---

## 3. 数据库-Excel交互功能

### 功能描述

在 `Datafeed` 类中新增的方法，用于将Excel文件导入数据库：
- 按目录树结构读取和处理Excel文件
- 错误检查和日志记录
- 保存错误文件
- 增量插入功能

### 主要方法

#### `get_existing_dates(code, metric)`
获取数据库中已有的日期序列
- **参数**:
  - `code`: 证券代码
  - `metric`: 指标名称
- **返回**: 日期字符串列表

#### `get_incremental_data(df, code_column='code', date_column='datetime', metric_column='metric')`
获取增量数据（数据库中不存在的数据）
- **参数**:
  - `df`: 待检查的DataFrame
  - `code_column`: 代码列名
  - `date_column`: 日期列名
  - `metric_column`: 指标列名
- **返回**: 增量数据DataFrame

#### `process_excel_to_db_format(filepath, transform_func=None, **kwargs)`
处理Excel文件为标准数据库格式
- **参数**:
  - `filepath`: Excel文件路径
  - `transform_func`: 自定义转换函数
  - `**kwargs`: 传递给转换函数的参数
- **返回**: (处理后的DataFrame, 检查结果字典)

#### `import_excel_to_db(filepath, transform_func=None, error_dir='./errors', validate_before_insert=True, incremental=False, **kwargs)`
将Excel文件导入到数据库
- **参数**:
  - `filepath`: Excel文件路径
  - `transform_func`: 数据转换函数
  - `error_dir`: 错误文件保存目录
  - `validate_before_insert`: 是否在插入前验证
  - `incremental`: 是否增量插入
- **返回**: 包含导入结果的字典

#### `batch_import_excel_to_db(directory, transform_func=None, pattern='*', error_dir='./errors', recursive=True, incremental=False, **kwargs)`
批量导入目录中的Excel文件到数据库
- **参数**:
  - `directory`: 目录路径
  - `transform_func`: 数据转换函数
  - `pattern`: 文件匹配模式
  - `error_dir`: 错误文件保存目录
  - `recursive`: 是否递归处理子目录
  - `incremental`: 是否增量插入
- **返回**: 包含成功和失败文件列表的字典

### 使用示例

```python
from betalens.datafeed import Datafeed
import pandas as pd

# 初始化Datafeed
datafeed = Datafeed('table_name')

# 定义转换函数
def transform_func(df):
    df['datetime'] = pd.to_datetime(df['date']) + pd.Timedelta(hours=15)
    df['metric'] = 'close_price'
    return df[['datetime', 'code', 'name', 'metric', 'value']]

# 导入单个文件
result = datafeed.import_excel_to_db(
    filepath='./data/stock_data.csv',
    transform_func=transform_func,
    incremental=True
)

# 批量导入
results = datafeed.batch_import_excel_to_db(
    directory='./data',
    transform_func=transform_func,
    pattern='*.csv',
    recursive=True,
    incremental=True
)

print(f"成功: {len(results['success'])} 个文件")
print(f"失败: {len(results['failed'])} 个文件")
```

---

## 4. 重构的查询方法

### 改进说明

重构了 `query_nearest_after()` 和 `query_nearest_before()` 方法，提高了代码的模块化和可维护性：

#### 新增辅助方法

##### `_build_time_filter(input_ts_col, data_ts_col, comparison, time_tolerance=None)`
构建时间过滤条件（解耦的辅助方法）
- **参数**:
  - `input_ts_col`: 输入时间戳列名
  - `data_ts_col`: 数据时间戳列名
  - `comparison`: 比较运算符 ('>' 或 '<=')
  - `time_tolerance`: 时间容差（小时）
- **返回**: SQL WHERE子句

##### `_execute_nearest_query(input_tuples, metric, order_direction='ASC', time_tolerance=None, filter_comparison='>')`
执行最近值查询的核心方法（解耦的辅助方法）
- **参数**:
  - `input_tuples`: (code, datetime)元组列表
  - `metric`: 指标名称
  - `order_direction`: 排序方向 'ASC'(之后) 或 'DESC'(之前)
  - `time_tolerance`: 时间容差（小时）
  - `filter_comparison`: 时间比较运算符
- **返回**: 查询结果DataFrame

### 改进优势

1. **更好的模块化**: 将共同逻辑提取到辅助方法中
2. **更易维护**: 代码更简洁，易于理解和修改
3. **更易扩展**: 可以轻松添加新的查询变体
4. **向后兼容**: 保持了原有API接口不变

### 使用示例

```python
from betalens.datafeed import Datafeed

# 初始化Datafeed
datafeed = Datafeed('daily_market_data')

# 查询之后最近的值
df_after = datafeed.query_nearest_after(params={
    'codes': ['000001.SZ', '000002.SZ'],
    'datetimes': ['2024-01-01 09:30:00', '2024-01-02 09:30:00'],
    'metric': 'close_price',
    'time_tolerance': 24  # 24小时
})

# 查询之前最近的值
df_before = datafeed.query_nearest_before(params={
    'codes': ['000001.SZ', '000002.SZ'],
    'datetimes': ['2024-01-01 09:30:00', '2024-01-02 09:30:00'],
    'metric': 'close_price',
    'time_tolerance': 24  # 24小时
})
```

---

## 日志和错误处理

### 日志记录

所有新增模块都包含完善的日志记录功能：
- **ExcelProcessor**: 自动记录文件读写操作、转换过程
- **DataValidator**: 记录所有验证警告和错误
- **数据库操作**: 记录导入过程和错误

### 错误处理

系统提供了多层错误处理机制：
1. **文件读取错误**: 自动捕获并记录
2. **数据验证错误**: 生成详细报告并保存
3. **数据库操作错误**: 自动回滚并保存错误文件
4. **错误文件保存**: 保存源文件、错误数据和错误日志

---

## 完整工作流示例

```python
from betalens.datafeed import ExcelProcessor, DataValidator, Datafeed
import pandas as pd

# 步骤1: 使用ExcelProcessor读取和处理文件
processor = ExcelProcessor(log_dir='./logs')
files = processor.batch_read_files('./raw_data', pattern='*.csv', recursive=True)

# 步骤2: 使用DataValidator验证数据质量
validator = DataValidator()
for filepath, df in files.items():
    # 检查数据质量
    check_results = validator.comprehensive_check(
        df,
        date_columns=['date'],
        required_columns=['code', 'date', 'value']
    )
    
    # 如果有问题，处理缺失值
    if check_results['null_values']:
        df = validator.handle_missing_values(df, method='ffill')
    
    # 转换为数据库格式
    db_df = processor.transform_to_db_format(
        df,
        key_columns=['code', 'name', 'date'],
        value_column='value',
        metric_name='close_price'
    )
    
    # 保存处理后的文件
    processor.save_dataframe(db_df, f'./processed/{filepath}')

# 步骤3: 使用Datafeed将数据导入数据库
datafeed = Datafeed('daily_market_data')

def transform_func(df):
    df['datetime'] = pd.to_datetime(df['date']) + pd.Timedelta(hours=15)
    return df

# 批量导入
results = datafeed.batch_import_excel_to_db(
    directory='./processed',
    transform_func=transform_func,
    pattern='*.csv',
    incremental=True,
    error_dir='./import_errors'
)

print(f"导入完成: 成功 {len(results['success'])} 个, 失败 {len(results['failed'])} 个")
```

---

## 注意事项

1. **数据库连接**: 确保PostgreSQL数据库正在运行且配置正确
2. **文件路径**: 建议使用绝对路径或相对于工作目录的路径
3. **数据格式**: 确保数据符合数据库表结构要求
4. **内存管理**: 处理大文件时注意内存使用
5. **错误日志**: 定期检查错误日志文件

---

## 版本信息

- **更新日期**: 2024-10-26
- **版本**: 2.0
- **作者**: Janis / GitHub Copilot

---

## 更多信息

详细示例请参考: `示例/excel_processing_demo.py`
