# data_service/box_service.py
import os
from typing import Optional, List
import geopandas as gpd
from constraints.field_name_mapper import (
    BOX_CODE_FIELD_NAME, BOX_CABLE_IN_FIELD_NAME,
    BOX_SKIP_COUNT_FIELD_NAME, BOX_IN_START_FIELD_NAME
)
from data_service.base import BaseLayerService


class BOXService(BaseLayerService):
    def __init__(self, gpkg_path, layer_name):
        # 初始化父类：指定BOX对应的GPKG和图层名

        super().__init__(gpkg_path=gpkg_path, layer_name=layer_name)

    def get_all_extremities(self):
        """
        获取所有终点类型的节点
        :return: 终点类型的节点列表
        """
        return self._dga.get_features_by_attribute('pass_seq', '==', 100)

    def boxs_amt_on_cable(self, cable_box):
        return self._dga.get_count_by_attribute("cable_in", "==", cable_box)

    def get_all_sro_points_by_orders(self, sort_by=None, ascending=True):
        """
        获取所有SRO节点，根据字段排序
        :param sort_by: 排序字段
        :param ascending: 是否升序 默认true
        :return: 所有SRO节点列表
        """
        return self._dga.get_features_by_attribute(field='class', op='==', value='SRO', sort_by=sort_by,
                                                   ascending=ascending)

    def get_box_by_code(self, box_code):
        boxes = self._dga.get_features_by_attribute(field=BOX_CODE_FIELD_NAME, op='==', value=box_code)
        if boxes.empty:
            return None
        else:
            return boxes.iloc[0]

    def get_all_sro_points_by_order_code_asc(self):
        """
        获取所有SRO节点，根据code升序
        :return: 所有SRO节点列表
        """
        return self.get_all_sro_points_by_orders(sort_by="code", ascending=True)

    def get_all_sro_points(self):
        """
        获取所有SRO节点
        :return: 所有SRO节点列表
        """
        return self.get_all_sro_points_by_orders()

    def get_all_points_on_cable_by_orders(self, cable_code: str, sort_by: Optional[list[str]] = None,
                                          ascending: bool | list[bool] = True):
        """
        获取指定线缆上所有掏芯节点（closure,终点）
        :param cable_code: 根据线缆code查询
        :param sort_by: 排序字段
        :param ascending: 是否升序 默认true
        :return: 所有掏芯节点列表
        """
        return self._dga.get_features_by_attribute(field='cable_in', op='==', value=cable_code, sort_by=sort_by,
                                                   ascending=ascending)

    def get_all_boxs_on_section_by_orders(self, section: int, sort_by: Optional[list[str]] = None,
                                          ascending: bool | list[bool] = True):
        """
        获取指定线缆上所有掏芯节点（closure,终点）
        :param section: 根据线缆code查询
        :param sort_by: 排序字段
        :param ascending: 是否升序 默认true
        :return: 所有掏芯节点列表
        """
        return self._dga.get_features_by_attribute(field=BOX_CABLE_IN_FIELD_NAME, op='==', value=section,
                                                   sort_by=sort_by,
                                                   ascending=ascending)

    def get_all_points_on_cable_by_order_in_start_asc(self, cable_code: str):
        """
        获取指定线缆上所有掏芯节点（closure,终点）
        :param cable_code: 根据线缆code查询
        :return: 所有掏芯节点列表
        """
        return self.get_all_points_on_cable_by_orders(cable_code=cable_code, sort_by=['in_start'], ascending=True)

    def get_all_points_on_cable(self, cable_code: str):
        """
        获取指定线缆上所有掏芯节点（closure,终点）
        :param cable_code: 根据线缆code查询
        :return: 所有掏芯节点列表
        """
        return self.get_all_points_on_cable_by_orders(cable_code=cable_code)

    def get_all_boxs_on_section(self, section: int):
        """
        获取指定线缆上所有掏芯节点（closure,终点）
        :param section: 根据线缆code查询
        :return: 所有掏芯节点列表
        """
        return self.get_all_boxs_on_section_by_orders(section=section)

    def init_data_of_all_sro_points(self):
        """
        初始化所有sro节点的skip_count值为0
        """

        def custom_condition(gdf):
            return gdf["class"] == 'SRO'

        update_success = self._dga.update_attributes(condition=custom_condition, field_values={'skip_count': 0})
        if update_success:
            self._dga.save_changes(overwrite=True)

    def update_skip_count_of_points_on_cable(self, cable_code: str, cable_skip_count: int):
        """
        更新指定线缆上所有掏芯节点的skip_count
        :param cable_code: 根据线缆code查询
        :param cable_skip_count: 线缆skip_count
        """
        print(f"cable_code: {cable_code}, skip count: {cable_skip_count}")

        def custom_condition(gdf):
            return gdf["cable_in"] == cable_code

        self._dga.update_attributes(condition=custom_condition,
                                    field_values={"skip_count": self._dga.gdf['in_start'] + cable_skip_count - 1})
        self._dga.save_changes(overwrite=True)

    def update_skip_count_of_boxs_on_section(self, section: int, section_skip_count: int):
        """
        更新指定线缆上所有掏芯节点的skip_count
        :param section: 根据线缆section查询
        :param section_skip_count: 线缆skip_count
        """
        print(f"section: {section}, skip count: {section_skip_count}")

        def custom_condition(gdf):
            return gdf[BOX_CABLE_IN_FIELD_NAME] == section

        self._dga.update_attributes(condition=custom_condition, field_values={
            BOX_SKIP_COUNT_FIELD_NAME: self._dga.gdf[BOX_IN_START_FIELD_NAME] + section_skip_count - 1})
        self._dga.save_changes(overwrite=True)
