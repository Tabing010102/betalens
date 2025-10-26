#%%By Janis 250226
# Todo：
'''
1. 单次存入csv数据
2. WIND API获取日行情并插入
3. 查询任意日期任意字段数据
研究数据库结构：
表1：个券行情数据，A股、A债、基金等的日频行情，列：入库实际时间（现实世界）=最早可用于交易时间（盘前中后）、windcode、中文名、数据性质（收盘价、成交量）、数值、备注（json，含说明等信息），同时计算close to close等收益率数据用于向量化回测和挖掘
    对于开盘价，在当天9点30出现。对于其他价量，最早15.00才可确定
表2：个券基本面数据，个券的财务等基本面，要对齐到券级别。列：入库实际时间（现实世界）=最早可用于交易时间（盘前中后）、数据理论上的发生时间（季报？）、windcode、中文名、数据性质（归母净利润）、数值、备注（json，含说明等信息）
    基本是日频，都发生在财报公布时间，注意公布时间在盘什么时间。理论上的发生时间标记统一到0331等，不关心未来函数，只作为画图展示等。
    如何处理年报次序公布：使用分析师一致预期、使用线性外推，尽量避免非原生数据的入库。
表3：宏观经济数据，入库实际时间（现实世界）=最早可用于交易时间（盘前中后）、数据理论上的发生时间（1月gdp）、windcode、中文名、数据性质（包含均线算子等）、数值、备注（json，含说明等信息）
    类似于表2.
表4：因子库数据，入库实际时间（现实世界）=最早可用于交易时间（盘前中后）、数据编制方式、数值、备注（json，含说明等信息）
投资数据库结构：
出于投资目的，实际不需要历史数据
从研究数据库拉取所需最近一个滚动窗口内的数据，其余全部在线拉取，并直接记录入库时间
'''
import numpy as np
import pandas as pd
import psycopg2
import psycopg2.extras
import os
import itertools
from psycopg2.extras import execute_values
from functools import wraps
import time
import logging
from pathlib import Path
from typing import Union, List, Dict, Optional, Tuple, Callable
import warnings

def func_timer(function):
    '''
    用装饰器实现函数计时
    :param function: 需要计时的函数
    :return: None
    '''
    @wraps(function)
    def function_timer(*args, **kwargs):
        print('[Function: {name} start...]'.format(name = function.__name__))
        t0 = time.time()
        result = function(*args, **kwargs)
        t1 = time.time()
        print('[Function: {name} finished, spent time: {time:.2f}s]'.format(name = function.__name__,time = t1 - t0))
        return result
    return function_timer


# ==================== Excel Processing Module ====================

class ExcelProcessor:
    """
    Excel文件处理模块
    功能：处理输入/输出的csv及xlsx文件，转化为pandas.dataframe
    """
    
    def __init__(self, log_dir: Optional[str] = None):
        """
        初始化Excel处理器
        
        Args:
            log_dir: 日志文件保存目录，如果为None则使用当前目录
        """
        self.log_dir = Path(log_dir) if log_dir else Path('./logs')
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # 配置日志
        self.logger = logging.getLogger('ExcelProcessor')
        self.logger.setLevel(logging.INFO)
        
        # 文件处理器
        log_file = self.log_dir / f'excel_processor_{time.strftime("%Y%m%d_%H%M%S")}.log'
        fh = logging.FileHandler(log_file, encoding='utf-8')
        fh.setLevel(logging.INFO)
        
        # 控制台处理器
        ch = logging.StreamHandler()
        ch.setLevel(logging.WARNING)
        
        # 格式化器
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        
        self.logger.addHandler(fh)
        self.logger.addHandler(ch)
    
    def read_file(self, filepath: Union[str, Path], **kwargs) -> Optional[pd.DataFrame]:
        """
        读取CSV或XLSX文件并转换为DataFrame
        
        Args:
            filepath: 文件路径
            **kwargs: 传递给pandas读取函数的额外参数
            
        Returns:
            DataFrame或None（如果读取失败）
        """
        filepath = Path(filepath)
        
        if not filepath.exists():
            self.logger.error(f"文件不存在: {filepath}")
            return None
        
        try:
            if filepath.suffix.lower() in ['.csv']:
                df = pd.read_csv(filepath, **kwargs)
                self.logger.info(f"成功读取CSV文件: {filepath}")
            elif filepath.suffix.lower() in ['.xlsx', '.xls']:
                df = pd.read_excel(filepath, **kwargs)
                self.logger.info(f"成功读取Excel文件: {filepath}")
            else:
                self.logger.error(f"不支持的文件格式: {filepath.suffix}")
                return None
            
            return df
            
        except Exception as e:
            self.logger.error(f"读取文件失败 {filepath}: {str(e)}")
            return None
    
    def save_dataframe(self, df: pd.DataFrame, filepath: Union[str, Path], 
                      file_format: str = 'csv', **kwargs) -> bool:
        """
        保存DataFrame到文件
        
        Args:
            df: 要保存的DataFrame
            filepath: 保存路径
            file_format: 文件格式 ('csv' 或 'xlsx')
            **kwargs: 传递给pandas保存函数的额外参数
            
        Returns:
            是否成功保存
        """
        filepath = Path(filepath)
        filepath.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            if file_format == 'csv':
                df.to_csv(filepath, encoding='utf_8_sig', index=False, **kwargs)
            elif file_format == 'xlsx':
                df.to_excel(filepath, index=False, **kwargs)
            else:
                self.logger.error(f"不支持的文件格式: {file_format}")
                return False
            
            self.logger.info(f"成功保存文件: {filepath}")
            return True
            
        except Exception as e:
            self.logger.error(f"保存文件失败 {filepath}: {str(e)}")
            return False
    
    def transform_to_db_format(self, df: pd.DataFrame, 
                               key_columns: List[str],
                               value_column: str,
                               metric_name: Optional[str] = None) -> pd.DataFrame:
        """
        将cross-section数据转化为数据库三列表格式
        
        Args:
            df: 输入DataFrame
            key_columns: 作为key的列名列表（如['code', 'name', 'date']）
            value_column: 包含值的列名
            metric_name: metric名称，如果为None则使用value_column的名称
            
        Returns:
            转换后的DataFrame，格式为 [key1, key2, ..., metric, value]
        """
        try:
            # 检查列是否存在
            missing_cols = set(key_columns + [value_column]) - set(df.columns)
            if missing_cols:
                self.logger.error(f"缺少列: {missing_cols}")
                return pd.DataFrame()
            
            # 选择需要的列
            result_df = df[key_columns + [value_column]].copy()
            
            # 添加metric列
            if metric_name:
                result_df['metric'] = metric_name
            else:
                result_df['metric'] = value_column
            
            # 重命名value列
            result_df = result_df.rename(columns={value_column: 'value'})
            
            self.logger.info(f"成功转换数据格式，行数: {len(result_df)}")
            return result_df
            
        except Exception as e:
            self.logger.error(f"转换数据格式失败: {str(e)}")
            return pd.DataFrame()
    
    def batch_read_files(self, directory: Union[str, Path], 
                        pattern: str = '*',
                        recursive: bool = False,
                        **kwargs) -> Dict[str, pd.DataFrame]:
        """
        批量读取目录中的文件
        
        Args:
            directory: 目录路径
            pattern: 文件匹配模式（如 '*.csv', '*.xlsx'）
            recursive: 是否递归读取子目录
            **kwargs: 传递给read_file的额外参数
            
        Returns:
            字典，键为文件路径，值为DataFrame
        """
        directory = Path(directory)
        results = {}
        
        if not directory.exists():
            self.logger.error(f"目录不存在: {directory}")
            return results
        
        # 查找文件
        if recursive:
            files = directory.rglob(pattern)
        else:
            files = directory.glob(pattern)
        
        for filepath in files:
            if filepath.is_file() and filepath.suffix.lower() in ['.csv', '.xlsx', '.xls']:
                df = self.read_file(filepath, **kwargs)
                if df is not None:
                    results[str(filepath)] = df
        
        self.logger.info(f"批量读取完成，成功读取 {len(results)} 个文件")
        return results
    
    def organize_output(self, dataframes: Dict[str, pd.DataFrame], 
                       output_dir: Union[str, Path],
                       categorize_by: Optional[Callable] = None) -> Dict[str, List[str]]:
        """
        批量保存文件并按分类组织
        
        Args:
            dataframes: 字典，键为原始文件名，值为DataFrame
            output_dir: 输出目录
            categorize_by: 分类函数，接收文件名返回分类名称
            
        Returns:
            目录树结构字典
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        directory_tree = {}
        
        for filename, df in dataframes.items():
            # 确定分类
            if categorize_by:
                category = categorize_by(filename)
            else:
                category = 'default'
            
            # 创建分类目录
            category_dir = output_dir / category
            category_dir.mkdir(parents=True, exist_ok=True)
            
            # 保存文件
            output_path = category_dir / Path(filename).name
            self.save_dataframe(df, output_path)
            
            # 记录目录树
            if category not in directory_tree:
                directory_tree[category] = []
            directory_tree[category].append(str(output_path))
        
        self.logger.info(f"文件组织完成，分类数: {len(directory_tree)}")
        return directory_tree


# ==================== Data Validation Utilities ====================

class DataValidator:
    """
    数据验证工具类
    用于检查DataFrame中的异常值、空值、格式错误等
    """
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """
        初始化数据验证器
        
        Args:
            logger: 日志记录器，如果为None则创建新的
        """
        if logger:
            self.logger = logger
        else:
            self.logger = logging.getLogger('DataValidator')
            self.logger.setLevel(logging.INFO)
            ch = logging.StreamHandler()
            ch.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            ch.setFormatter(formatter)
            self.logger.addHandler(ch)
    
    def check_null_values(self, df: pd.DataFrame, 
                         columns: Optional[List[str]] = None) -> Dict[str, int]:
        """
        检查DataFrame中的空值、NaN、None值
        
        Args:
            df: 待检查的DataFrame
            columns: 要检查的列名列表，如果为None则检查所有列
            
        Returns:
            字典，键为列名，值为空值数量
        """
        if columns is None:
            columns = df.columns.tolist()
        
        null_counts = {}
        
        for col in columns:
            if col not in df.columns:
                self.logger.warning(f"列 '{col}' 不存在")
                continue
            
            # 检查多种空值类型
            null_count = df[col].isna().sum() + df[col].isnull().sum()
            none_count = (df[col] == None).sum()
            
            total_null = null_count + none_count
            null_counts[col] = int(total_null)
            
            if total_null > 0:
                self.logger.warning(f"列 '{col}' 包含 {total_null} 个空值")
        
        return null_counts
    
    def check_date_column(self, df: pd.DataFrame, 
                         date_column: str,
                         check_duplicates: bool = True,
                         check_sorting: bool = True,
                         check_frequency: bool = False,
                         expected_freq: Optional[str] = None) -> Dict[str, any]:
        """
        检查日期列的各种问题
        
        Args:
            df: 待检查的DataFrame
            date_column: 日期列名
            check_duplicates: 是否检查重复日期
            check_sorting: 是否检查排序
            check_frequency: 是否检查频率
            expected_freq: 期望的频率（如 'D', 'W', 'M'）
            
        Returns:
            包含检查结果的字典
        """
        results = {
            'has_duplicates': False,
            'is_sorted': True,
            'frequency_consistent': True,
            'issues': []
        }
        
        if date_column not in df.columns:
            self.logger.error(f"日期列 '{date_column}' 不存在")
            results['issues'].append(f"列不存在: {date_column}")
            return results
        
        # 尝试转换为datetime
        try:
            dates = pd.to_datetime(df[date_column])
        except Exception as e:
            self.logger.error(f"日期列 '{date_column}' 包含无效的日期格式: {str(e)}")
            results['issues'].append(f"日期格式错误: {str(e)}")
            return results
        
        # 检查重复
        if check_duplicates:
            duplicates = dates.duplicated()
            if duplicates.any():
                dup_count = duplicates.sum()
                results['has_duplicates'] = True
                results['issues'].append(f"发现 {dup_count} 个重复日期")
                self.logger.warning(f"日期列 '{date_column}' 包含 {dup_count} 个重复日期")
        
        # 检查排序
        if check_sorting:
            is_sorted = dates.is_monotonic_increasing
            results['is_sorted'] = is_sorted
            if not is_sorted:
                results['issues'].append("日期未按升序排列")
                self.logger.warning(f"日期列 '{date_column}' 未按升序排列")
        
        # 检查频率
        if check_frequency and len(dates) > 1:
            try:
                inferred_freq = pd.infer_freq(dates.sort_values())
                if expected_freq and inferred_freq != expected_freq:
                    results['frequency_consistent'] = False
                    results['issues'].append(f"频率不一致，期望: {expected_freq}, 实际: {inferred_freq}")
                    self.logger.warning(f"日期频率不一致，期望: {expected_freq}, 实际: {inferred_freq}")
            except Exception as e:
                results['frequency_consistent'] = False
                results['issues'].append(f"无法推断日期频率: {str(e)}")
                self.logger.warning(f"无法推断日期频率: {str(e)}")
        
        return results
    
    def check_data_types(self, df: pd.DataFrame, 
                        expected_types: Dict[str, type]) -> Dict[str, bool]:
        """
        检查列的数据类型是否符合预期
        
        Args:
            df: 待检查的DataFrame
            expected_types: 字典，键为列名，值为期望的类型
            
        Returns:
            字典，键为列名，值为是否匹配
        """
        results = {}
        
        for col, expected_type in expected_types.items():
            if col not in df.columns:
                self.logger.warning(f"列 '{col}' 不存在")
                results[col] = False
                continue
            
            # 检查是否为数值型列
            if expected_type in [int, float, np.number]:
                is_valid = pd.api.types.is_numeric_dtype(df[col])
            elif expected_type == str:
                is_valid = pd.api.types.is_string_dtype(df[col]) or df[col].dtype == 'object'
            elif expected_type in [pd.Timestamp, np.datetime64]:
                is_valid = pd.api.types.is_datetime64_any_dtype(df[col])
            else:
                is_valid = df[col].dtype == expected_type
            
            results[col] = is_valid
            
            if not is_valid:
                self.logger.warning(f"列 '{col}' 的数据类型不匹配，期望: {expected_type}, 实际: {df[col].dtype}")
        
        return results
    
    def handle_missing_values(self, df: pd.DataFrame, 
                            method: str = 'drop',
                            columns: Optional[List[str]] = None,
                            fill_value: any = None,
                            **kwargs) -> pd.DataFrame:
        """
        处理缺失值
        
        Args:
            df: 待处理的DataFrame
            method: 处理方法 ('drop', 'ffill', 'bfill', 'fill', 'raise')
            columns: 要处理的列，如果为None则处理所有列
            fill_value: 当method='fill'时使用的填充值
            **kwargs: 传递给处理方法的额外参数
            
        Returns:
            处理后的DataFrame
        """
        df_copy = df.copy()
        
        if columns:
            subset = columns
        else:
            subset = df_copy.columns.tolist()
        
        if method == 'drop':
            df_copy = df_copy.dropna(subset=subset, **kwargs)
            self.logger.info(f"删除包含空值的行，剩余 {len(df_copy)} 行")
        
        elif method == 'ffill':
            df_copy[subset] = df_copy[subset].ffill(**kwargs)
            self.logger.info("使用前向填充处理空值")
        
        elif method == 'bfill':
            df_copy[subset] = df_copy[subset].bfill(**kwargs)
            self.logger.info("使用后向填充处理空值")
        
        elif method == 'fill':
            if fill_value is None:
                self.logger.error("使用fill方法时必须提供fill_value参数")
                return df
            df_copy[subset] = df_copy[subset].fillna(fill_value, **kwargs)
            self.logger.info(f"使用值 {fill_value} 填充空值")
        
        elif method == 'raise':
            null_counts = self.check_null_values(df_copy, columns=subset)
            total_nulls = sum(null_counts.values())
            if total_nulls > 0:
                raise ValueError(f"发现 {total_nulls} 个空值，列: {null_counts}")
        
        else:
            self.logger.error(f"不支持的处理方法: {method}")
            return df
        
        return df_copy
    
    def comprehensive_check(self, df: pd.DataFrame, 
                          date_columns: Optional[List[str]] = None,
                          required_columns: Optional[List[str]] = None) -> Dict[str, any]:
        """
        综合检查DataFrame
        
        Args:
            df: 待检查的DataFrame
            date_columns: 日期列列表
            required_columns: 必须存在的列
            
        Returns:
            包含所有检查结果的字典
        """
        results = {
            'shape': df.shape,
            'null_values': {},
            'date_checks': {},
            'missing_columns': [],
            'errors': []
        }
        
        # 检查必需列
        if required_columns:
            missing = set(required_columns) - set(df.columns)
            if missing:
                results['missing_columns'] = list(missing)
                results['errors'].append(f"缺少必需列: {missing}")
        
        # 检查空值
        results['null_values'] = self.check_null_values(df)
        
        # 检查日期列
        if date_columns:
            for date_col in date_columns:
                if date_col in df.columns:
                    results['date_checks'][date_col] = self.check_date_column(df, date_col)
        
        # 检查空行
        empty_rows = df.isnull().all(axis=1).sum()
        if empty_rows > 0:
            results['errors'].append(f"发现 {empty_rows} 个空行")
            self.logger.warning(f"发现 {empty_rows} 个空行")
        
        return results

class Datafeed():
    _initialized = False

    def __init__(self, sheetname):
      #  if not self._initialized:
            self.conn = psycopg2.connect(
                dbname="datafeed",
                user="postgres",
                password="111111",#你自己的卡密
                host="localhost",
                port="5432"
            )
            self.cursor = self.conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            self.sheet = sheetname
            self.__class__._initialized = True
            
            # 初始化辅助工具
            self.excel_processor = ExcelProcessor()
            self.data_validator = DataValidator()

    # ==================== Database-Excel Interaction Module ====================
    
    def get_existing_dates(self, code: str, metric: str) -> List[str]:
        """
        获取数据库中已有的日期序列
        
        Args:
            code: 证券代码
            metric: 指标名称
            
        Returns:
            日期字符串列表
        """
        query = f"""
            SELECT DISTINCT datetime 
            FROM {self.sheet} 
            WHERE code = %s AND metric = %s 
            ORDER BY datetime
        """
        try:
            self.cursor.execute(query, (code, metric))
            dates = [str(row['datetime']) for row in self.cursor.fetchall()]
            return dates
        except Exception as e:
            print(f"获取已有日期失败: {str(e)}")
            return []
    
    def get_incremental_data(self, df: pd.DataFrame, 
                            code_column: str = 'code',
                            date_column: str = 'datetime',
                            metric_column: str = 'metric') -> pd.DataFrame:
        """
        获取增量数据（数据库中不存在的数据）
        
        Args:
            df: 待检查的DataFrame
            code_column: 代码列名
            date_column: 日期列名
            metric_column: 指标列名
            
        Returns:
            增量数据DataFrame
        """
        if df.empty:
            return df
        
        # 确保日期列是datetime类型
        df_copy = df.copy()
        df_copy[date_column] = pd.to_datetime(df_copy[date_column])
        
        # 获取所有唯一的code和metric组合
        unique_combinations = df_copy[[code_column, metric_column]].drop_duplicates()
        
        # 标记哪些行是新数据
        new_data_mask = pd.Series([False] * len(df_copy))
        
        for _, row in unique_combinations.iterrows():
            code = row[code_column]
            metric = row[metric_column]
            
            # 获取数据库中已有的日期
            existing_dates = self.get_existing_dates(code, metric)
            existing_dates_set = set(pd.to_datetime(existing_dates))
            
            # 找出新数据
            mask = (df_copy[code_column] == code) & (df_copy[metric_column] == metric)
            df_dates = df_copy.loc[mask, date_column]
            is_new = ~df_dates.isin(existing_dates_set)
            new_data_mask[mask] = is_new.values
        
        return df_copy[new_data_mask]
    
    def process_excel_to_db_format(self, filepath: Union[str, Path],
                                   transform_func: Optional[Callable] = None,
                                   **kwargs) -> Tuple[Optional[pd.DataFrame], Dict]:
        """
        处理Excel文件为标准数据库格式
        
        Args:
            filepath: Excel文件路径
            transform_func: 自定义转换函数
            **kwargs: 传递给转换函数的参数
            
        Returns:
            (处理后的DataFrame, 检查结果字典)
        """
        # 读取文件
        df = self.excel_processor.read_file(filepath)
        if df is None:
            return None, {'error': '文件读取失败'}
        
        # 应用转换函数
        if transform_func:
            try:
                df = transform_func(df, **kwargs)
            except Exception as e:
                return None, {'error': f'转换失败: {str(e)}'}
        
        # 数据验证
        check_results = self.data_validator.comprehensive_check(
            df,
            date_columns=['datetime'] if 'datetime' in df.columns else None,
            required_columns=['code', 'metric', 'value'] if transform_func else None
        )
        
        return df, check_results
    
    def import_excel_to_db(self, filepath: Union[str, Path],
                          transform_func: Optional[Callable] = None,
                          error_dir: Optional[str] = './errors',
                          validate_before_insert: bool = True,
                          incremental: bool = False,
                          **kwargs) -> Dict[str, any]:
        """
        将Excel文件导入到数据库
        
        Args:
            filepath: Excel文件路径
            transform_func: 数据转换函数
            error_dir: 错误文件保存目录
            validate_before_insert: 是否在插入前验证
            incremental: 是否增量插入（只插入新数据）
            **kwargs: 传递给转换函数的参数
            
        Returns:
            包含导入结果的字典
        """
        result = {
            'success': False,
            'filepath': str(filepath),
            'rows_processed': 0,
            'rows_inserted': 0,
            'errors': []
        }
        
        # 处理文件
        df, check_results = self.process_excel_to_db_format(filepath, transform_func, **kwargs)
        
        if df is None:
            result['errors'].append(check_results.get('error', '未知错误'))
            self._save_error_file(filepath, error_dir, check_results)
            return result
        
        result['rows_processed'] = len(df)
        
        # 检查是否有错误
        if validate_before_insert and check_results.get('errors'):
            result['errors'].extend(check_results['errors'])
            self._save_error_file(filepath, error_dir, check_results, df)
            return result
        
        # 增量插入
        if incremental:
            df = self.get_incremental_data(df)
            if df.empty:
                result['success'] = True
                result['rows_inserted'] = 0
                print(f"文件 {filepath} 没有新数据需要插入")
                return result
        
        # 插入数据
        try:
            insert_result = self.insert_washed_data(df, self.sheet)
            if insert_result == 0:
                result['success'] = True
                result['rows_inserted'] = len(df)
                print(f"成功插入 {len(df)} 行数据")
            else:
                result['errors'].append('数据库插入失败')
        except Exception as e:
            result['errors'].append(f'插入异常: {str(e)}')
            self._save_error_file(filepath, error_dir, {'error': str(e)}, df)
        
        return result
    
    def batch_import_excel_to_db(self, directory: Union[str, Path],
                                 transform_func: Optional[Callable] = None,
                                 pattern: str = '*',
                                 error_dir: Optional[str] = './errors',
                                 recursive: bool = True,
                                 incremental: bool = False,
                                 **kwargs) -> Dict[str, List]:
        """
        批量导入目录中的Excel文件到数据库
        
        Args:
            directory: 目录路径
            transform_func: 数据转换函数
            pattern: 文件匹配模式
            error_dir: 错误文件保存目录
            recursive: 是否递归处理子目录
            incremental: 是否增量插入
            **kwargs: 传递给转换函数的参数
            
        Returns:
            包含成功和失败文件列表的字典
        """
        directory = Path(directory)
        results = {
            'success': [],
            'failed': [],
            'total_rows': 0,
            'total_inserted': 0
        }
        
        if not directory.exists():
            print(f"目录不存在: {directory}")
            return results
        
        # 查找所有文件
        if recursive:
            files = list(directory.rglob(pattern))
        else:
            files = list(directory.glob(pattern))
        
        # 过滤出Excel文件
        excel_files = [f for f in files if f.suffix.lower() in ['.csv', '.xlsx', '.xls']]
        
        print(f"找到 {len(excel_files)} 个文件待处理")
        
        # 处理每个文件
        for filepath in excel_files:
            print(f"\n处理文件: {filepath}")
            result = self.import_excel_to_db(
                filepath,
                transform_func=transform_func,
                error_dir=error_dir,
                incremental=incremental,
                **kwargs
            )
            
            results['total_rows'] += result['rows_processed']
            results['total_inserted'] += result['rows_inserted']
            
            if result['success']:
                results['success'].append(str(filepath))
            else:
                results['failed'].append({
                    'filepath': str(filepath),
                    'errors': result['errors']
                })
        
        print(f"\n批量导入完成:")
        print(f"  成功: {len(results['success'])} 个文件")
        print(f"  失败: {len(results['failed'])} 个文件")
        print(f"  总行数: {results['total_rows']}")
        print(f"  插入行数: {results['total_inserted']}")
        
        return results
    
    def _save_error_file(self, filepath: Union[str, Path],
                        error_dir: str,
                        check_results: Dict,
                        df: Optional[pd.DataFrame] = None):
        """
        保存错误文件和日志
        
        Args:
            filepath: 源文件路径
            error_dir: 错误文件保存目录
            check_results: 检查结果
            df: 错误数据DataFrame（如果有）
        """
        error_dir = Path(error_dir)
        error_dir.mkdir(parents=True, exist_ok=True)
        
        filepath = Path(filepath)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        
        # 保存源文件副本
        error_file = error_dir / f"error_{timestamp}_{filepath.name}"
        try:
            import shutil
            shutil.copy2(filepath, error_file)
        except Exception as e:
            print(f"保存错误源文件失败: {str(e)}")
        
        # 保存错误数据（如果有）
        if df is not None:
            data_file = error_dir / f"error_data_{timestamp}_{filepath.stem}.csv"
            df.to_csv(data_file, encoding='utf_8_sig', index=False)
        
        # 保存错误日志
        log_file = error_dir / f"error_log_{timestamp}_{filepath.stem}.txt"
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f"文件: {filepath}\n")
            f.write(f"时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"检查结果:\n")
            import json
            f.write(json.dumps(check_results, ensure_ascii=False, indent=2))
        
        print(f"错误文件已保存到: {error_dir}")
    
    # ==================== Refactored Query Helper Methods ====================
    
    def _build_time_filter(self, input_ts_col: str, data_ts_col: str, 
                          comparison: str, time_tolerance: Optional[float] = None) -> str:
        """
        构建时间过滤条件（解耦的辅助方法）
        
        Args:
            input_ts_col: 输入时间戳列名
            data_ts_col: 数据时间戳列名
            comparison: 比较运算符 ('>' 或 '<=')
            time_tolerance: 时间容差（小时）
            
        Returns:
            SQL WHERE子句
        """
        filter_clause = f"{data_ts_col} {comparison} {input_ts_col}"
        
        if time_tolerance:
            time_diff = f"({data_ts_col} - {input_ts_col})" if comparison == '>' else f"({input_ts_col} - {data_ts_col})"
            filter_clause += f" AND {time_diff} <= {time_tolerance} * INTERVAL '1 hour'"
        
        return filter_clause
    
    def _execute_nearest_query(self, input_tuples: List[Tuple], 
                              metric: str,
                              order_direction: str = 'ASC',
                              time_tolerance: Optional[float] = None,
                              filter_comparison: str = '>') -> pd.DataFrame:
        """
        执行最近值查询的核心方法（解耦的辅助方法）
        
        Args:
            input_tuples: (code, datetime)元组列表
            metric: 指标名称
            order_direction: 排序方向 'ASC'(之后) 或 'DESC'(之前)
            time_tolerance: 时间容差（小时）
            filter_comparison: 时间比较运算符
            
        Returns:
            查询结果DataFrame
        """
        # 生成占位符
        value_placeholders = ', '.join(['(%s, %s::TIMESTAMP)'] * len(input_tuples))
        
        # 构建时间过滤条件
        time_filter = self._build_time_filter('i.input_ts', 't.datetime', 
                                              filter_comparison, time_tolerance)
        
        # 构建SQL
        sql = f"""WITH input_data (code, input_ts) AS (
            VALUES {value_placeholders}
        ),
        candidate_data AS (
            SELECT
                i.code,
                i.input_ts,
                t.datetime AS datetime,
                EXTRACT(EPOCH FROM (t.datetime - i.input_ts))/3600 AS diff_hours,
                t.value,
                ROW_NUMBER() OVER (
                    PARTITION BY i.code, i.input_ts 
                    ORDER BY t.datetime {order_direction}
                ) AS rn
            FROM input_data i
            LEFT JOIN {self.sheet} t
                ON i.code = t.code
                AND {time_filter}
                AND t.metric = %s
        )
        SELECT 
            code,
            input_ts,
            datetime,
            diff_hours,
            value
        FROM candidate_data
        WHERE rn = 1
        """
        
        # 构造参数
        params_list = []
        for code, dt in input_tuples:
            params_list.extend([code, dt])
        params_list.append(metric)
        
        # 执行查询
        self.cursor.execute(sql, params_list)
        return pd.DataFrame(self.cursor.fetchall())

    @staticmethod
    def insert_daily_market_data(self, data, table):

        data = pd.melt(data, id_vars=data.columns[0:3], value_vars=data.columns[3:-1])
        data.loc[data['variable'] == "开盘价(元)", '日期'] += " 09:30:01"
        data.loc[data['variable'] != "开盘价(元)", '日期'] += " 15:00:01"
        data['日期'] = pd.to_datetime(data['日期'])

        df = data[['日期','代码','简称','variable','value']]
        df.columns = ['datetime', 'code', 'name', 'metric', 'value']

        import psycopg2.extras as extras
        tuples = [tuple(x) for x in df.to_numpy()]

        cols = ','.join(list(df.columns))
        # SQL query to execute
        query = "INSERT INTO %s(%s) VALUES %%s" % (table, cols)
        try:
            extras.execute_values(self.cursor, query, tuples)
            self.conn.commit()
        except (Exception, psycopg2.DatabaseError) as error:
            print("Error: %s" % error)
            self.conn.rollback()
            self.cursor.close()
            return 1
        #print("the dataframe is inserted")
        #self.cursor.close()
        return 0

    @staticmethod
    def insert_daily_index_data(self, data, table):

        data = pd.melt(data, id_vars=data.columns[0:3], value_vars=data.columns[3:-1])
        data.loc[data['variable'] == "开盘价", '日期'] += " 09:30:01"
        data.loc[data['variable'] != "开盘价", '日期'] += " 15:00:01"
        data['日期'] = pd.to_datetime(data['日期'])

        df = data[['日期', '代码', '简称', 'variable', 'value']]
        df.columns = ['datetime', 'code', 'name', 'metric', 'value']

        import psycopg2.extras as extras
        tuples = [tuple(x) for x in df.to_numpy()]

        cols = ','.join(list(df.columns))
        # SQL query to execute
        query = "INSERT INTO %s(%s) VALUES %%s" % (table, cols)
        try:
            extras.execute_values(self.cursor, query, tuples)
            self.conn.commit()
        except (Exception, psycopg2.DatabaseError) as error:
            print("Error: %s" % error)
            self.conn.rollback()
            self.cursor.close()
            return 1
        # print("the dataframe is inserted")
        # self.cursor.close()
        return 0


    @staticmethod
    def insert_daily_fund_data(self, data, table):

        data = pd.melt(data, id_vars=data.columns[0:3], value_vars=data.columns[3:-1])
        data.loc[data['variable'] == "开盘价(元)", '日期'] += " 09:30:01"
        data.loc[data['variable'] != "开盘价(元)", '日期'] += " 15:00:01"
        data['日期'] = pd.to_datetime(data['日期'])

        df = data[['日期', '代码', '简称', 'variable', 'value']]
        df.columns = ['datetime', 'code', 'name', 'metric', 'value']

        import psycopg2.extras as extras
        tuples = [tuple(x) for x in df.to_numpy()]

        cols = ','.join(list(df.columns))
        # SQL query to execute
        query = "INSERT INTO %s(%s) VALUES %%s" % (table, cols)
        try:
            extras.execute_values(self.cursor, query, tuples)
            self.conn.commit()
        except (Exception, psycopg2.DatabaseError) as error:
            print("Error: %s" % error)
            self.conn.rollback()
            self.cursor.close()
            return 1
        # print("the dataframe is inserted")
        # self.cursor.close()
        return 0

    @staticmethod
    def wash_ede_data(filepath):
        def text_to_json(input_text):
            # 去除输入文本两端的双引号（如果存在）
            import json
            cleaned_text = input_text.strip('"')

            # 按行分割文本
            lines = cleaned_text.split('\n')

            # 创建结果字典
            result = {
                lines[0]: {}
            }

            # 处理属性行
            for line in lines[1:]:
                if line.strip():  # 跳过空行
                    # 分割键值对
                    parts = line.split(']', 1)  # 只在第一个']'处分割

                    # 提取键名（去除开头的'['）
                    key = parts[0].replace('[', '')

                    # 提取值并去除多余空格
                    value_str = parts[1].strip() if len(parts) > 1 else ""

                    # 尝试将字符串值转换为合适的类型
                    try:
                        # 尝试转换为整数
                        value = int(value_str)
                    except ValueError:
                        try:
                            # 尝试转换为浮点数
                            value = float(value_str)
                        except ValueError:
                            # 保留为字符串
                            value = value_str

                    # 添加到结果字典
                    result[lines[0]][key] = value

            return json.dumps(result, ensure_ascii=False, indent=2)

        df = pd.read_csv(filepath)
        df.rename(columns={'证券代码': 'code', '证券简称': 'name', }, inplace=True)
        df = pd.melt(df, id_vars=['code', 'name'], var_name='note', value_name='value')

        import re
        try:
            df['label_datetime'] = df['note'].apply(lambda x :re.search(r'\[报表年度\]\s*(\d{4})', x).group(1))
        except:
            df['label_datetime'] = df['note'].apply(lambda x: re.search(r'\[报告期\]\s*(\d{4})', x).group(1))
        try:
            df['note'] = df['note'].apply(text_to_json)
        except:
            print("note转写json失败")
        df.to_csv(filepath+"_washed.csv", encoding='utf_8_sig', index=False)
        return df

    @staticmethod
    def merge_fundamental_data(label_time_filepath, value_filepath, metric_name):
        label_time = pd.read_csv(label_time_filepath)
        value = pd.read_csv(value_filepath)

        label_time.rename(columns={'value':'datetime',}, inplace=True)
        label_time['datetime'] += " 15:00:01"
        df = pd.merge(label_time[['code', 'name', 'datetime', 'label_datetime']], value, on=['code', 'name', 'label_datetime'])
        df['metric'] = metric_name
        df.to_csv(value_filepath+"_mergeded.csv", encoding='utf_8_sig',index=False)

        return 0

    def insert_washed_data(self, df, table):

        import psycopg2.extras as extras
        tuples = [tuple(x) for x in df.to_numpy()]

        cols = ','.join(list(df.columns))
        # SQL query to execute
        query = "INSERT INTO %s(%s) VALUES %%s" % (table, cols)
        try:
            extras.execute_values(self.cursor, query, tuples)
            self.conn.commit()
        except (Exception, psycopg2.DatabaseError) as error:
            print("Error: %s" % error)
            self.conn.rollback()
            self.cursor.close()
            return 1
        # print("the dataframe is inserted")
        # self.cursor.close()
        return 0


    def insert_files(self, folder_path, insert_func):
        # 获取文件夹中所有文件的列表
        file_list = os.listdir(folder_path)
        error_list = []
        # 读取所有CSV文件并将其存储在一个列表中
        for file in file_list:
            if file.endswith('.CSV') or file.endswith('.csv') :
                file_path = os.path.join(folder_path, file)
                try:
                    data = pd.read_csv(file_path)
                except:
                    print(file_path + " Error!")
                    error_list.append(file_path)
                    continue
                if(insert_func(self, data, self.sheet)):
                    print(file_path + " Error!")
                    error_list.append(file_path)
                else:
                    print(file_path+" inserted successfully")

        return error_list



    def check_result(self, result:pd.DataFrame):
        result.drop("note", axis=1, inplace=True)
        result.drop_duplicates(inplace=True)
        result.sort_values(by=['datetime'], inplace=True)
        result.reset_index(drop=True, inplace=True)

        def check(group):
            # 检查 datetime 是否有重复值
            duplicate_dates = group['datetime'].duplicated(keep=False)
            if duplicate_dates.any():
                print("Warning: 发现重复的 datetime 值！")
                # 统计每个重复的日期出现的次数
                date_counts = group['datetime'].value_counts()
                for date, count in date_counts.items():
                    if count > 1:
                        print(f"日期 {date} 出现了 {count} 次。")

        result.groupby(['code', 'metric']).apply(check, include_groups=False)
        return result

    @func_timer
    def query_data(self, params=None):
        """
        查询 daily_market_data 表中的数据

        Args:
            conn: 数据库连接
            cursor: 数据库游标
            params (dict): 包含查询条件的字典，支持以下键：
                - 'start_date': 开始日期，格式如 '2023-01-01'
                - 'end_date': 结束日期，格式如 '2024-01-01'
                - 'code': 要查询的代码（股票代码等）[]
                - 'metric': 要查询的指标

        Returns:
            pandas DataFrame: 查询结果
        """
        conditions = []
        params_list = []

        if params is None:
            params = {}

        # 遍历字典中的每个键值对
        for key, value in params.items():
            if value is not None:
                if key == 'start_date':
                    conditions.append("datetime >= '" + value + "' :: TIMESTAMP")
                    params_list.append(value)
                elif key == 'end_date':
                    conditions.append("datetime <= '" + value + "' :: TIMESTAMP")
                    params_list.append(value)
                elif key == 'code':
                    conditions.append("(code = '" + "' OR code = '".join(value) + "')")
                    params_list.append(value)
                elif key == 'metric':
                    conditions.append("metric = '" + value + "'")
                    params_list.append(value)
                elif key == 'label_start_date':
                    conditions.append("label_datetime >= '" + value + "' :: TIMESTAMP")
                    params_list.append(value)
                elif key == 'label_end_date':
                    conditions.append("label_datetime <= '" + value + "' :: TIMESTAMP")
                    params_list.append(value)

        # 构建查询语句
        query = "SELECT * FROM " + self.sheet
        if conditions:
            query += " WHERE " + " AND ".join(conditions)

        # 执行查询
        print(query)
        self.cursor.execute(query)

        # 将结果转换为 DataFrame
        result = self.check_result(pd.DataFrame(self.cursor.fetchall()))

        return result

    #时间结构：最新特征<=提数时点<调仓时点   
    #时间戳：提数时点：提数时点：晚于提数时点的价格时点
    @func_timer
    def query_nearest_after(self, params=None):
        """
        根据输入时间戳序列查找每个时点之后最近的有效值
        主要用于回测时提取价格，
        
        重构说明：使用_execute_nearest_query辅助方法实现更好的模块化和解耦
        
        Args:
            params (dict): 必须包含以下键：
                - codes: 代码列表（必须与datetimes等长）
                - datetimes: 目标时间戳列表（格式：'YYYY-MM-DD HH:MM'）
                - metric: 查询的指标名称
                - time_tolerance: 允许的最大时间间隔（单位：小时，默认不限制）

        Returns:
            DataFrame: 包含以下列：
                code | input_datetime | matched_datetime | time_diff_hours | value
        """
        # 参数校验
        required_keys = ['codes', 'datetimes', 'metric']
        if not all(k in params for k in required_keys):
            raise ValueError(f"必须提供参数: {required_keys}")

        # 生成输入数据对
        def gen_pairs():
            for code, dt in itertools.product(params['codes'], params['datetimes']):
                yield (code, dt)

        input_tuples = list(gen_pairs())
        
        # 使用重构后的辅助方法执行查询
        df = self._execute_nearest_query(
            input_tuples=input_tuples,
            metric=params['metric'],
            order_direction='ASC',  # 找之后最近的，按时间升序
            time_tolerance=params.get('time_tolerance'),
            filter_comparison='>'  # 找之后的数据
        )
        
        df.rename(columns={'value': params['metric']}, inplace=True)
        return df

    @func_timer
    def query_nearest_before(self, params=None):
        """
        根据输入时间戳序列查找每个时点之前最近的有效值
        主要用于回测时提取历史价格特征，时间结构需满足：
            调仓时点 <= 提数时点 < 最新特征时点
        
        重构说明：使用_execute_nearest_query辅助方法实现更好的模块化和解耦
        
        Args:
            params (dict): 必须包含以下键：
                - codes: 代码列表（必须与datetimes等长）
                - datetimes: 目标时间戳列表（格式：'YYYY-MM-DD HH:MM'）
                - metric: 查询的指标名称
                - time_tolerance: 允许的最大时间间隔（单位：小时，默认不限制）

        Returns:
            DataFrame: 包含以下列：
                code | input_datetime | matched_datetime | time_diff_hours | value
        """
        # 参数校验
        required_keys = ['codes', 'datetimes', 'metric']
        if not all(k in params for k in required_keys):
            raise ValueError(f"必须提供参数: {required_keys}")

        # 生成输入数据对
        def gen_pairs():
            for code, dt in itertools.product(params['codes'], params['datetimes']):
                yield (code, dt)

        input_tuples = list(gen_pairs())
        
        # 使用重构后的辅助方法执行查询
        df = self._execute_nearest_query(
            input_tuples=input_tuples,
            metric=params['metric'],
            order_direction='DESC',  # 找之前最近的，按时间降序
            time_tolerance=params.get('time_tolerance'),
            filter_comparison='<='  # 找之前的数据
        )
        
        df.rename(columns={'value': params['metric']}, inplace=True)
        return df

def get_absolute_trade_days(begin_date, end_date, period):
    #string
    from WindPy import w
    w.start()
    trade_days = w.tdays(begin_date, end_date, "Period="+period).Data[0]
    return trade_days


#%%
# if __name__ == '__main__':

    # data.wash_ede_data(r"C:\Users\Janis\OneDrive\因子框架\股息率(报告期).csv")
    # data.wash_ede_data(r"C:\Users\Janis\OneDrive\因子框架\定期报告实际披露日期.csv")
    # data.merge_fundamental_data(r"C:\Users\Janis\OneDrive\因子框架\定期报告实际披露日期.csv"+"_washed.csv",r"C:\Users\Janis\OneDrive\因子框架\股息率(报告期).csv"+"_washed.csv","股息率(报告期)")
    # error_result = data.insert_files(r"C:\Users\Janis\OneDrive\因子框架\insert", data.insert_washed_data)
    # d = Datafeed("fundamental_data")
    # os.chdir(r"C:\Users\Janis\Desktop")
    # data = pd.read_pickle("交易状态.pkl")
    # data['metric'] = "交易状态"
    # import datetime as dt
    # data['datetime'] = data['note'] + dt.timedelta(hours=9,minutes=32)
    # data['label_datetime'] = data['note']
    # codes, uniques = pd.factorize(data['value'])
    # data['note'] = "{\"交易状态\":\"" + data['value'] + "\"}"
    # data['value'] = codes
    # data.to_pickle("交易状态95-25.pkl")
    # d.insert_washed_data(df=data, table=d.sheet)
    # 示例调用
    # insert_fundamental_data('annual_report.csv', 'dividend_yield.csv', 'fundamental_data')
    # error_result = db.insert_files(r"C:\Users\Janis\Desktop\基金", db.insert_daily_fund_data)
    # pd.DataFrame(error_result).to_excel("error_result2.xlsx")
