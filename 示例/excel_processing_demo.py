"""
Excel处理模块和数据验证工具示例

演示如何使用新增的功能：
1. Excel处理模块 (ExcelProcessor)
2. 数据验证工具 (DataValidator)
3. 数据库-Excel交互功能 (Datafeed的新方法)
"""

import pandas as pd
import numpy as np
import sys
import os

# 添加父目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from betalens.datafeed import ExcelProcessor, DataValidator, Datafeed


def demo_excel_processor():
    """演示Excel处理器的使用"""
    print("=" * 60)
    print("演示 1: Excel处理器 (ExcelProcessor)")
    print("=" * 60)
    
    # 初始化处理器
    processor = ExcelProcessor(log_dir='./logs')
    
    # 示例1: 创建测试数据
    test_data = pd.DataFrame({
        'code': ['000001.SZ', '000002.SZ', '000001.SZ', '000002.SZ'],
        'name': ['平安银行', '万科A', '平安银行', '万科A'],
        'date': ['2024-01-01', '2024-01-01', '2024-01-02', '2024-01-02'],
        'close_price': [10.5, 8.3, 10.8, 8.5]
    })
    
    print("\n原始数据:")
    print(test_data)
    
    # 示例2: 转换为数据库三列表格式
    print("\n转换为数据库格式...")
    db_format = processor.transform_to_db_format(
        test_data,
        key_columns=['code', 'name', 'date'],
        value_column='close_price',
        metric_name='收盘价'
    )
    
    print("\n数据库格式:")
    print(db_format)
    
    # 示例3: 保存文件
    # processor.save_dataframe(db_format, './output/test_data.csv', file_format='csv')
    print("\n✓ Excel处理器演示完成")


def demo_data_validator():
    """演示数据验证器的使用"""
    print("\n" + "=" * 60)
    print("演示 2: 数据验证工具 (DataValidator)")
    print("=" * 60)
    
    # 初始化验证器
    validator = DataValidator()
    
    # 创建包含各种问题的测试数据
    test_data = pd.DataFrame({
        'code': ['000001.SZ', '000002.SZ', None, '000003.SZ', '000001.SZ'],
        'datetime': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04', '2024-01-02'],
        'value': [10.5, np.nan, 8.3, 9.1, 10.8],
        'metric': ['close', 'close', 'close', 'close', 'close']
    })
    
    print("\n测试数据（包含问题）:")
    print(test_data)
    
    # 示例1: 检查空值
    print("\n检查空值...")
    null_counts = validator.check_null_values(test_data)
    print(f"空值统计: {null_counts}")
    
    # 示例2: 检查日期列
    print("\n检查日期列...")
    date_check = validator.check_date_column(
        test_data,
        'datetime',
        check_duplicates=True,
        check_sorting=True
    )
    print(f"日期检查结果: {date_check}")
    
    # 示例3: 处理缺失值 - 前向填充
    print("\n使用前向填充处理空值...")
    filled_data = validator.handle_missing_values(
        test_data,
        method='ffill',
        columns=['value']
    )
    print(filled_data)
    
    # 示例4: 综合检查
    print("\n执行综合检查...")
    check_results = validator.comprehensive_check(
        test_data,
        date_columns=['datetime'],
        required_columns=['code', 'datetime', 'value', 'metric']
    )
    print(f"综合检查结果:")
    print(f"  数据形状: {check_results['shape']}")
    print(f"  空值: {check_results['null_values']}")
    print(f"  错误: {check_results['errors']}")
    
    print("\n✓ 数据验证器演示完成")


def demo_database_excel_interaction():
    """演示数据库-Excel交互功能"""
    print("\n" + "=" * 60)
    print("演示 3: 数据库-Excel交互功能")
    print("=" * 60)
    
    print("\n注意: 此演示需要数据库连接，仅展示使用方法")
    print("实际使用时需要配置数据库连接参数")
    
    # 示例代码（需要数据库连接）
    print("\n示例代码:")
    print("""
# 初始化Datafeed（需要数据库连接）
# datafeed = Datafeed('table_name')

# 定义转换函数
def transform_func(df):
    # 将原始数据转换为标准格式
    df['datetime'] = pd.to_datetime(df['date']) + pd.Timedelta(hours=15)
    df['metric'] = 'close_price'
    return df[['datetime', 'code', 'name', 'metric', 'value']]

# 示例1: 导入单个Excel文件
# result = datafeed.import_excel_to_db(
#     filepath='./data/stock_data.csv',
#     transform_func=transform_func,
#     incremental=True  # 增量插入
# )
# print(f"导入结果: {result}")

# 示例2: 批量导入目录中的文件
# results = datafeed.batch_import_excel_to_db(
#     directory='./data',
#     transform_func=transform_func,
#     pattern='*.csv',
#     recursive=True,
#     incremental=True
# )
# print(f"成功: {len(results['success'])} 个文件")
# print(f"失败: {len(results['failed'])} 个文件")

# 示例3: 使用重构后的查询方法
# df = datafeed.query_nearest_after(params={
#     'codes': ['000001.SZ', '000002.SZ'],
#     'datetimes': ['2024-01-01 09:30:00', '2024-01-02 09:30:00'],
#     'metric': 'close_price',
#     'time_tolerance': 24  # 24小时
# })
# print(df)
    """)
    
    print("\n✓ 数据库-Excel交互功能演示完成")


def demo_refactored_query_methods():
    """演示重构后的查询方法"""
    print("\n" + "=" * 60)
    print("演示 4: 重构后的查询方法")
    print("=" * 60)
    
    print("\n查询方法的改进:")
    print("1. query_nearest_after() - 使用辅助方法实现更好的模块化")
    print("2. query_nearest_before() - 使用辅助方法实现更好的模块化")
    print("3. 增加了以下辅助方法:")
    print("   - _build_time_filter(): 构建时间过滤条件")
    print("   - _execute_nearest_query(): 执行最近值查询")
    print("\n这些改进使代码更易维护和扩展，同时保持了向后兼容性")
    
    print("\n✓ 查询方法重构演示完成")


if __name__ == '__main__':
    print("\n")
    print("*" * 60)
    print("*" + " " * 58 + "*")
    print("*" + "  Excel处理模块和数据验证工具示例".center(56) + "*")
    print("*" + " " * 58 + "*")
    print("*" * 60)
    
    # 运行所有演示
    demo_excel_processor()
    demo_data_validator()
    demo_database_excel_interaction()
    demo_refactored_query_methods()
    
    print("\n" + "=" * 60)
    print("所有演示完成！")
    print("=" * 60)
    print("\n主要功能总结:")
    print("1. ExcelProcessor - 处理Excel/CSV文件，支持批量操作和格式转换")
    print("2. DataValidator - 验证数据质量，检查空值、日期等问题")
    print("3. 数据库-Excel交互 - 将Excel数据导入数据库，支持增量更新")
    print("4. 重构的查询方法 - 更模块化、更易维护的数据库查询功能")
    print()
