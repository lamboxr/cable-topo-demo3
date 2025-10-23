import geopandas as gpd
from typing import Optional, Dict, Callable, List

import numpy as np
from geopandas import GeoDataFrame
from shapely.geometry import base
import pandas as pd


class LayerDGA:
    """按图层标识创建单例的GeoPackage图层操作类"""
    _instances: Dict[str, "LayerDGA"] = {}  # 存储{图层标识: 实例}的映射
    gpkg_path: str
    layer_name: str
    _gdf: Optional[gpd.GeoDataFrame]

    def __new__(cls, gpkg_path: str, layer_name: str):
        """
        按图层标识（gpkg_path + layer_name）创建单例
        :param gpkg_path: GeoPackage文件路径
        :param layer_name: 图层名称
        """
        # 生成唯一标识（区分不同图层）
        instance_key = f"{gpkg_path}::{layer_name}"
        if instance_key not in cls._instances:
            # 新标识：创建实例并缓存
            cls._instances[instance_key] = super().__new__(cls)
            # 初始化实例属性
            cls._instances[instance_key].gpkg_path = gpkg_path
            cls._instances[instance_key].layer_name = layer_name
            cls._instances[instance_key]._gdf = None  # 缓存当前图层的GeoDataFrame
            # 加载图层数据
            cls._instances[instance_key]._load_layer()
        return cls._instances[instance_key]

    def _load_layer(self):
        """加载当前图层数据到缓存"""
        try:
            self._gdf = gpd.read_file(self.gpkg_path, layer=self.layer_name)
            print(f"图层加载成功：{self.layer_name}@{self.gpkg_path}，要素数：{len(self._gdf)}")
        except Exception as e:
            print(f"图层加载失败：{self.layer_name}@{self.gpkg_path}，错误：{str(e)}")
            self._gdf = None

    @property
    def gdf(self) -> Optional[gpd.GeoDataFrame]:
        """获取当前图层的GeoDataFrame（自动重新加载如果未加载）"""
        if self._gdf is None:
            self._load_layer()
        return self._gdf

    # --------------------------
    # 新增：数据更新方法
    # --------------------------
    # def update_attributes(self, condition: Callable, field: str, new_value) -> bool:
    #     """
    #     按条件更新指定字段的值
    #     :param condition: 筛选条件（如：lambda gdf: gdf["level"] == 1）
    #     :param field: 要更新的字段名
    #     :param new_value: 新值（可是固定值或与字段长度匹配的列表/数组）
    #     :return: 更新成功返回True，失败返回False
    #     """
    #     if self.gdf is None or field not in self.gdf.columns:
    #         print(f"图层为空或字段不存在：{field}")
    #         return False
    #
    #     # 筛选符合条件的行并更新字段值
    #     mask = condition(self.gdf)
    #     if not mask.any():
    #         print("无符合条件的要素可更新")
    #         return True
    #
    #     self._gdf.loc[mask, field] = new_value
    #     print(f"已更新 {mask.sum()} 个要素的 {field} 字段")
    #     return True

    # def update_attributes(self,
    #                       condition: Callable,
    #                       field: str,
    #                       new_value,
    #                       value_processor: Optional[Callable] = None) -> bool:
    #     """
    #     按条件更新指定字段的值（支持处理空值/非数字值）
    #     :param condition: 筛选条件（如：lambda gdf: gdf["level"] == 1）
    #     :param field: 要更新的字段名（A字段）
    #     :param new_value: 新值表达式（如：gdf["B"] + 10）
    #     :param value_processor: 值处理器（可选），用于清洗new_value中的无效值
    #                             （如：lambda x: x.fillna(0) 填充空值为0）
    #     :return: 更新成功返回True
    #     """
    #     if self.gdf is None or field not in self.gdf.columns:
    #         print(f"图层为空或字段不存在：{field}")
    #         return False
    #
    #     # 筛选符合条件的行
    #     mask = condition(self.gdf)
    #     if not mask.any():
    #         print("无符合条件的要素可更新")
    #         return True
    #
    #     # 计算原始新值（可能包含空值/非数字）
    #     raw_new_value = new_value
    #
    #     # 应用值处理器（清洗无效值）
    #     if value_processor is not None:
    #         processed_new_value = value_processor(raw_new_value)
    #     else:
    #         processed_new_value = raw_new_value
    #
    #     # 仅更新符合条件的行
    #     self._gdf.loc[mask, field] = processed_new_value
    #
    #     # 统计有效更新数量（排除因无效值导致的未更新）
    #     if isinstance(processed_new_value, (pd.Series, np.ndarray)):
    #         # 数组类型：用 isna() 判断空值
    #         valid_mask = mask & ~processed_new_value.isna()
    #     else:
    #         # 单个值（如 int/float）：非空即有效
    #         valid_mask = mask  # 单个值无空值问题，直接使用原始条件掩码
    #     print(f"已更新 {self.layer_name} 图层 {valid_mask.sum()} 个要素的 {field} 字段（总符合条件 {mask.sum()} 个）")
    #     return True
    def update_attributes(
            self,
            condition: Optional[Callable] = None,  # 支持为空（更新全部）
            field_values: Dict[str, any] = None,  # 多字段更新：键=字段名，值=新值/表达式
            value_processor: Optional[Callable] = None  # 支持对单个字段传处理器（可选）
    ) -> bool:
        """
        按条件更新指定字段的值（支持condition为空、多字段更新、处理无效值）
        :param condition: 筛选条件（可为None，此时更新全部要素）
        :param field_values: 多字段更新字典（如：{"A": gdf["B"]+10, "C": 0}）
        :param value_processor: 值处理器（可选），支持两种形式：
                                1. 单个函数：对所有字段的新值统一处理
                                2. 字典：对不同字段传不同处理器（如：{"A": lambda x: x.fillna(0), "C": lambda x: x.astype(int)}）
        :return: 更新成功返回True，失败返回False
        """
        # 1. 基础校验
        if self.gdf is None:
            print("图层为空，无法更新")
            return False
        if field_values is None or not isinstance(field_values, dict):
            print("field_values必须是非空字典（键=字段名，值=新值）")
            return False
        # 校验所有字段是否存在
        invalid_fields = [field for field in field_values.keys() if field not in self.gdf.columns]
        if invalid_fields:
            print(f"以下字段不存在：{invalid_fields}，无法更新")
            return False

        # 2. 处理condition为空：生成全为True的掩码（更新全部）
        if condition is None:
            mask = pd.Series([True] * len(self.gdf), index=self.gdf.index, dtype=bool)
            print("condition为空，将更新图层所有要素")
        else:
            # 执行条件函数并校验结果
            try:
                mask = condition(self.gdf)
                if mask is None:
                    raise ValueError("条件函数返回None，预期为布尔类型Series")
                if not isinstance(mask, pd.Series) or mask.dtype != bool:
                    raise TypeError("条件函数返回值必须是布尔类型Series")
            except Exception as e:
                print(f"条件函数执行失败：{str(e)}")
                return False

        # 3. 校验符合条件的要素（无符合条件时直接返回）
        if not mask.any():
            print("无符合条件的要素可更新")
            return True

        # 4. 遍历多字段，逐个更新
        total_valid = 0  # 统计所有字段的有效更新总数
        for field, new_value in field_values.items():
            # 4.1 计算当前字段的原始新值
            raw_new_value = new_value

            # 4.2 选择当前字段的处理器（优先字段专属，无则用全局）
            if isinstance(value_processor, dict):
                field_processor = value_processor.get(field)  # 字段专属处理器
            else:
                field_processor = value_processor  # 全局处理器

            # 4.3 应用处理器清洗新值
            if field_processor is not None:
                try:
                    processed_new_value = field_processor(raw_new_value)
                except Exception as e:
                    print(f"字段 {field} 的值处理器执行失败：{str(e)}，跳过该字段更新")
                    continue
            else:
                processed_new_value = raw_new_value

            # 4.4 执行字段更新（仅更新符合条件的行）
            try:
                self._gdf.loc[mask, field] = processed_new_value
            except Exception as e:
                print(f"字段 {field} 更新失败：{str(e)}，跳过该字段更新")
                continue

            # 4.5 统计当前字段的有效更新数量
            if isinstance(processed_new_value, (pd.Series, np.ndarray, gpd.GeoSeries)):
                # 数组类型：排除空值
                valid_mask = mask & ~processed_new_value.isna()
            else:
                # 单个值：所有符合条件的行均有效
                valid_mask = mask
            field_valid_count = valid_mask.sum()
            total_valid += field_valid_count
            print(
                f"已更新 {self.layer_name} 图层 {field_valid_count} 个要素的 {field} 字段（该字段符合条件 {mask.sum()} 个）")

        # 5. 最终结果提示
        if total_valid == 0:
            print("所有字段均未完成有效更新")
        else:
            print(f"多字段更新完成，累计有效更新 {total_valid} 次（含多个字段叠加）")
        return True

    def add_features(self, new_features: gpd.GeoDataFrame) -> bool:
        """
        向图层添加新要素
        :param new_features: 包含新要素的GeoDataFrame（需与原图层字段和CRS一致）
        :return: 添加成功返回True
        """
        if self.gdf is None:
            print("图层未加载，无法添加要素")
            return False

        # 检查字段一致性（新要素必须包含原图层的所有必填字段）
        missing_fields = set(self.gdf.columns) - set(new_features.columns)
        if missing_fields:
            print(f"新要素缺少必要字段：{missing_fields}")
            return False

        # 检查坐标系一致性
        if self.gdf.crs != new_features.crs:
            print(f"坐标系不一致，自动转换为图层坐标系：{self.gdf.crs}")
            new_features = new_features.to_crs(self.gdf.crs)

        # 拼接新要素到原图层
        self._gdf = pd.concat([self._gdf, new_features], ignore_index=True)
        print(f"已添加 {len(new_features)} 个新要素，当前总要素数：{len(self._gdf)}")
        return True

    def delete_features(self, condition: Callable) -> bool:
        """
        按条件删除要素
        :param condition: 筛选条件（如：lambda gdf: gdf["status"] == "废弃"）
        :return: 删除成功返回True
        """
        if self.gdf is None:
            print("图层未加载，无法删除要素")
            return False

        # 筛选符合条件的行
        mask = condition(self.gdf)
        if not mask.any():
            print("无符合条件的要素可删除")
            return True

        # 删除要素（保留不满足条件的要素）
        self._gdf = self._gdf[~mask].copy()
        print(f"已删除 {mask.sum()} 个要素，当前总要素数：{len(self._gdf)}")
        return True

    def save_changes(self, overwrite: bool = True) -> bool:
        """
        将内存中的修改同步到GeoPackage文件
        :param overwrite: 是否覆盖原图层
        :return: 保存成功返回True
        """
        if self.gdf is None:
            print("无数据可保存")
            return False

        try:
            # 保存图层（覆盖模式）
            self._gdf.to_file(
                self.gpkg_path,
                layer=self.layer_name,
                driver="GPKG",
                mode="w",
                append=not overwrite
            )
            print(f"修改已同步到文件：{self.gpkg_path}（图层：{self.layer_name}）")
            return True
        except Exception as e:
            print(f"保存失败：{e}")
            return False

    # --------------------------
    # 新增：按条件查询要素集合
    # --------------------------
    def get_features_by_condition(
            self,
            condition: Optional[Callable[[gpd.GeoDataFrame], bool]] = None,  # 允许为None
            sort_by: Optional[list[str]] = None,  # 改为列表：支持多个字段（如["level", "voltage"]）
            ascending: bool | list[bool] = True  # 改为列表/单个布尔值：对应每个字段的排序方向
    ) -> Optional[gpd.GeoDataFrame]:
        """
        根据条件查询要素集合，支持按多个字段排序
        :param condition: 筛选条件（如：lambda gdf: gdf["level"] == 1）
        :param sort_by: 排序字段列表（如：["level", "voltage"]，None表示不排序）
        :param ascending: 排序方向（单个布尔值或列表）：
                          - 若为单个值：所有字段使用同一方向（True=升序，False=降序）
                          - 若为列表：需与sort_by长度一致，分别指定每个字段的方向
        :return: 筛选并排序后的GeoDataFrame
        """
        if self.gdf is None:
            print("图层数据为空，无法查询")
            return None
        # try:
        # 1. 处理condition为None的情况：返回全部要素（布尔掩码全为True）
        if condition is None:
            condition_result = pd.Series([True] * len(self.gdf), index=self.gdf.index, dtype=bool)
        else:
            # 执行条件函数
            condition_result = condition(self.gdf)

            # 2. 检查返回值是否为None
            if condition_result is None:
                raise ValueError("条件函数返回了None，预期应为布尔类型的Series（布尔掩码）")

            # 3. 检查返回值是否为有效的布尔掩码（pandas.Series且dtype为bool）
            if not isinstance(condition_result, pd.Series):
                raise TypeError(f"条件函数返回值类型错误，预期为pd.Series，实际为{type(condition_result)}")
            if condition_result.dtype != bool:
                raise TypeError(f"条件函数返回值应为布尔类型（bool），实际为{condition_result.dtype}")

        # 4. 筛选符合条件的要素
        filtered_gdf = self.gdf[condition_result].copy()
        if filtered_gdf.empty:
            print("无符合条件的要素")
            return filtered_gdf

        # 2. 处理排序逻辑
        if sort_by is not None:
            # 确保sort_by是列表（兼容单个字段传入的情况）
            if isinstance(sort_by, str):
                sort_by = [sort_by]

            # 检查所有排序字段是否存在
            missing_fields = [f for f in sort_by if f not in filtered_gdf.columns]
            if missing_fields:
                print(f"排序字段不存在：{missing_fields}，将跳过排序")
                return filtered_gdf

            # 处理排序方向（确保与sort_by长度一致）
            if isinstance(ascending, bool):
                # 单个布尔值：所有字段使用同一方向
                ascending = [ascending] * len(sort_by)
            elif len(ascending) != len(sort_by):
                print(f"排序方向列表长度（{len(ascending)}）与字段数量（{len(sort_by)}）不匹配，将跳过排序")
                return filtered_gdf

            # 执行多字段排序
            filtered_gdf = filtered_gdf.sort_values(
                by=sort_by,
                ascending=ascending,
                na_position="last"  # 空值放在最后
            )
            # 打印排序信息
            sort_info = ", ".join([
                f"{field} {'升序' if asc else '降序'}"
                for field, asc in zip(sort_by, ascending)
            ])
            print(f"已按以下规则排序：{sort_info}")

        print(f"查询到 {len(filtered_gdf)} 个符合条件的要素")
        return filtered_gdf
        # except Exception as e:
        #     print(f"查询失败：{str(e)}")
        #     return gpd.GeoDataFrame()


    # 便捷方法：属性查询（简化常用场景）
    def get_features_by_attribute(
            self,
            field: str,
            op: str,
            value,
            sort_by: Optional[list[str]] = None,  # 多字段排序
            ascending: bool | list[bool] = True
    ) -> Optional[gpd.GeoDataFrame]:
        """
        按属性条件查询（如：field="voltage", op=">", value=10）
        :param field: 字段名
        :param op: 运算符（">", "<", "==", "contains"等）
        :param value: 比较值
        :param sort_by: 排序字段列表（如：["level", "voltage"]，None表示不排序）
        :param ascending: 排序方向（单个布尔值或列表）：
                          - 若为单个值：所有字段使用同一方向（True=升序，False=降序）
                          - 若为列表：需与sort_by长度一致，分别指定每个字段的方向
        :return: 符合条件的要素集合
        """
        if self.gdf is None or field not in self.gdf.columns:
            print(f"字段不存在或图层为空：{field}")
            return None

        def custom_condition(gdf):
            return gen_condition(gdf, field, op, value)

        return self.get_features_by_condition(custom_condition, sort_by, ascending)

    def get_count_by_condition(
            self,
            condition: Optional[Callable[[gpd.GeoDataFrame], bool]] = None,  # 允许为None
    ) -> Optional[gpd.GeoDataFrame]:
        """
        根据条件查询要素集合，支持按多个字段排序
        :param condition: 筛选条件（如：lambda gdf: gdf["level"] == 1）
        :return: 筛选并排序后的GeoDataFrame
        """
        if self.gdf is None:
            print("图层数据为空，无法查询")
            return None
        # try:
        # 1. 处理condition为None的情况：返回全部要素（布尔掩码全为True）
        if condition is None:
            condition_result = pd.Series([True] * len(self.gdf), index=self.gdf.index, dtype=bool)
        else:
            # 执行条件函数
            condition_result = condition(self.gdf)

            # 2. 检查返回值是否为None
            if condition_result is None:
                raise ValueError("条件函数返回了None，预期应为布尔类型的Series（布尔掩码）")

            # 3. 检查返回值是否为有效的布尔掩码（pandas.Series且dtype为bool）
            if not isinstance(condition_result, pd.Series):
                raise TypeError(f"条件函数返回值类型错误，预期为pd.Series，实际为{type(condition_result)}")
            if condition_result.dtype != bool:
                raise TypeError(f"条件函数返回值应为布尔类型（bool），实际为{condition_result.dtype}")

        # 4. 筛选符合条件的要素
        filtered_gdf = self.gdf[condition_result].copy()
        return condition_result.sum()




    def get_count_by_attribute(
            self,
            field: str,
            op: str,
            value
    ) -> int | GeoDataFrame | None:
        """
        按属性条件统计要素数量（如：field="voltage", op=">", value=10）
        :param field: 字段名
        :param op: 运算符（">", "<", "==", "contains"等）
        :param value: 比较值
        :return: 符合条件的要素数
        """
        if self.gdf is None or field not in self.gdf.columns:
            print(f"图层为空或字段不存在：{field}，计数为0")
            return -1

        # 构建属性条件（复用查询方法的逻辑）

        try:
            def custom_condition(gdf):
                return gen_condition(gdf, field, op, value)

            if custom_condition is None:
                return -1
            else:
                return self.get_count_by_condition(custom_condition)


        except Exception as e:
            print(f"属性计数失败：{str(e)}")
            return -1

    # 便捷方法：空间查询（简化常用场景）
    def get_features_by_spatial(self, spatial_op: str, geometry) -> Optional[gpd.GeoDataFrame]:
        """
        按空间条件查询（如：包含、相交、距离小于等）
        :param spatial_op: 空间运算符（"contains", "intersects", "distance_lt"等）
        :param geometry: 参考几何对象（如Point, Polygon）
        :return: 符合条件的要素集合
        """
        if self.gdf is None:
            print("图层数据为空，无法查询")
            return None

        # 确保几何对象与图层坐标系一致
        if self.gdf.crs != geometry.crs:
            geometry = geometry.to_crs(self.gdf.crs)

        # 构建空间条件
        if spatial_op == "contains":
            # 图层要素包含参考几何
            condition = self.gdf.geometry.contains(geometry)
        elif spatial_op == "intersects":
            # 图层要素与参考几何相交
            condition = self.gdf.geometry.intersects(geometry)
        elif spatial_op == "distance_lt":
            # 图层要素与参考几何的距离小于指定值（需米制坐标系）
            distance = geometry  # 此时geometry参数传入距离值（如1000米）
            condition = self.gdf.geometry.distance(geometry) < distance
        else:
            print(f"不支持的空间运算符：{spatial_op}")
            return None

        filtered_gdf = self.gdf[condition].copy()
        print(f"空间查询到 {len(filtered_gdf)} 个要素")
        return filtered_gdf

    # --------------------------
    # 旧的查询方法
    # --------------------------
    def query_by_attribute(self, condition) -> Optional[gpd.GeoDataFrame]:
        """属性查询（返回符合条件的子集）"""
        if self.gdf is None:
            return None
        return self.gdf[condition].copy()

    # --------------------------
    # 旧的查询方法
    # --------------------------
    def query_by_spatial(self, spatial_func) -> Optional[gpd.GeoDataFrame]:
        """空间查询（返回符合空间条件的子集）"""
        if self.gdf is None:
            return None
        return self.gdf[spatial_func(self.gdf)].copy()

    # --------------------------
    # 旧的保存方法
    # --------------------------
    def save(self, overwrite: bool = True):
        """保存当前图层（覆盖原图层）"""
        if self.gdf is None:
            print("无数据可保存")
            return
        # 保存图层（覆盖模式）
        self.gdf.to_file(
            self.gpkg_path,
            layer=self.layer_name,
            driver="GPKG",
            mode="a",
            append=not overwrite  # overwrite=True时append=False（覆盖）
        )
        print(f"图层已保存：{self.layer_name}@{self.gpkg_path}")

    def refresh(self):
        """重新加载图层数据（刷新缓存）"""
        self._load_layer()


def gen_condition(gdf, field, op, value):
    """生成条件表达式，确保value为标量"""
    # 关键：如果value是Series，尝试提取第一个值（或根据业务逻辑取标量）
    if isinstance(value, pd.Series):
        # 若Series非空，取第一个值；否则返回空条件
        if not value.empty:
            value = value.iloc[0]  # 提取第一个元素作为标量
        else:
            return pd.Series([False] * len(gdf), index=gdf.index)  # 无匹配值
    if op == ">":
        return gdf[field] > value
    elif op == "<":
        return gdf[field] < value
    elif op == "==":
        return gdf[field] == value
    elif op == "!=":
        return gdf[field] != value
    elif op == "contains":
        # 文本包含（需字段为字符串类型）
        return gdf[field].astype(str).str.contains(str(value))
    elif op == "in":
        # 检查是否在列表中（如：value=[1,2,3]）
        return gdf[field].isin(value)
    else:
        print(f"不支持的运算符：{op}")
        return None


# 使用示例
if __name__ == "__main__":
    # 1. 创建cable图层的单例实例
    gpkg_cable_path = "/gpkg/cable.gpkg"
    cable_dga1 = LayerDGA(gpkg_cable_path, "cable")
    cable_dga2 = LayerDGA(gpkg_cable_path, "cable")  # 同一图层，应返回同一实例
    print(f"cable实例是否相同：{cable_dga1 is cable_dga2}")  # 输出：True

    gpkg_nap_path = "/gpkg/nap.gpkg"
    # 2. 创建nap图层的单例实例
    nap_dga1 = LayerDGA(gpkg_nap_path, "nap")
    nap_dga2 = LayerDGA(gpkg_nap_path, "nap")  # 同一图层，应返回同一实例
    print(f"nap实例是否相同：{nap_dga1 is nap_dga2}")  # 输出：True

    # 3. 验证cable和nap是不同实例
    print(f"cable和nap实例是否不同：{cable_dga1 is nap_dga1}")  # 输出：False

    # 4. 使用cable实例查询
    if cable_dga1.gdf is not None:
        # 查询电压>10kV的电缆
        level1 = cable_dga1.query_by_attribute(
            condition=cable_dga1.gdf["level"] == 1
        )
        print(f"cable图层中level==1的要素数：{len(level1)}")

    # 5. 使用nap实例查询
    if nap_dga1.gdf is not None:
        # 查询code包含"NAP"的点
        nap_filtered = nap_dga1.query_by_attribute(
            condition=nap_dga1.gdf["in_start"] == 1
        )
        print(f"nap图层中in_start==1的要素数：{len(nap_filtered)}")
