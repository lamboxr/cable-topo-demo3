# data_service/base_layer_service.py
import os
from typing import Optional, Callable, List, Dict
import geopandas as gpd
from utils.gda_utils import LayerDGA
from utils import gpkg_utils


class BaseLayerService:
    """图层服务基类，封装LayerDGA的基础操作"""

    def __init__(self, gpkg_path: str, layer_name: str):
        """
        初始化图层服务
        :param gpkg_name: GPKG文件名（如"SRO.gpkg"）
        :param layer_name: 图层名（如"elj_qae_sro"）
        """
        print(f"========os.path.exists({gpkg_path}) : {os.path.exists(gpkg_path)}")
        self._dga = LayerDGA(gpkg_path, layer_name)  # 关联LayerDGA实例

    @property
    def dga(self) -> LayerDGA:
        """获取LayerDGA实例（单例）"""
        return self._dga

    @property
    def gdf(self) -> Optional[gpd.GeoDataFrame]:
        """获取图层GeoDataFrame（简化调用）"""
        return self.dga.gdf

    # 封装基础查询方法（复用LayerDGA的能力）
    def get_by_condition(
            self,
            condition: Optional[Callable] = None,
            sort_by: Optional[List[str]] = None,
            ascending: bool | List[bool] = True
    ) -> Optional[gpd.GeoDataFrame]:
        """按条件查询要素（代理LayerDGA的get_features_by_condition）"""
        return self.dga.get_features_by_condition(
            condition=condition,
            sort_by=sort_by,
            ascending=ascending
        )

    def get_by_attribute(
            self,
            field: str,
            op: str,
            value,
            sort_by: Optional[List[str]] = None,
            ascending: bool | List[bool] = True
    ) -> Optional[gpd.GeoDataFrame]:
        """按属性查询要素（代理LayerDGA的get_features_by_attribute）"""
        return self.dga.get_features_by_attribute(
            field=field,
            op=op,
            value=value,
            sort_by=sort_by,
            ascending=ascending
        )

    def update_by_condition(
            self,
            condition: Optional[Callable] = None,
            field_values: Dict[str, any] = None,
            value_processor: Optional[Callable] = None
    ) -> bool:
        """按条件更新要素（代理LayerDGA的update_attributes）"""
        return self.dga.update_attributes(
            condition=condition,
            field_values=field_values,
            value_processor=value_processor
        )

    def save_changes(self, overwrite: bool = True) -> bool:
        """保存修改（代理LayerDGA的save_changes）"""
        return self.dga.save_changes(overwrite=overwrite)
