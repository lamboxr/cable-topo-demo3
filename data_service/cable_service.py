# data_service/cable_service.py
from typing import Optional, List
import geopandas as gpd
from constraints.field_name_mapper import (
    CABLE_ORIGIN_FIELD_NAME, CABLE_CODE_FIELD_NAME,
    CABLE_EXTREMITY_FIELD_NAME, CABLE_SECTION_FIELD_NAME,
    CABLE_LEVEL_FIELD_NAME, CABLE_ORIGIN_BOX_FIELD_NAME, CABLE_SKIP_COUNT_FIELD_NAME, CABLE_PORT_START_FIELD_NAME
)
from data_service.base import BaseLayerService


class CABLEService(BaseLayerService):
    def __init__(self, gpkg_path, layer_name):
        # 初始化父类：指定CABLE对应的GPKG和图层名
        super().__init__(gpkg_path=gpkg_path, layer_name=layer_name)

    def get_next_segment_by_origin_code(self, section_value, box_code):
        def custom_condition(gdf):
            return (gdf[CABLE_SECTION_FIELD_NAME] == section_value) & (
                    gdf[CABLE_ORIGIN_FIELD_NAME] == box_code)

        next_section = self._dga.get_features_by_condition(custom_condition)
        if next_section is None or next_section.empty:
            return None
        return next_section.iloc[0]

    def get_next_segment(self, segment):
        section = segment[CABLE_SECTION_FIELD_NAME]
        extremity = segment[CABLE_EXTREMITY_FIELD_NAME]
        return self.get_next_segment_by_origin_code(section, extremity)

    def get_all_1st_segments_on_d1_section_order_by_skip_count_asc(self):
        def custom_condition(gdf):
            return (gdf[CABLE_LEVEL_FIELD_NAME] == 1) & (
                        gdf[CABLE_ORIGIN_FIELD_NAME] == gdf[CABLE_ORIGIN_BOX_FIELD_NAME])

        return self._dga.get_features_by_condition(custom_condition, sort_by=[CABLE_SKIP_COUNT_FIELD_NAME])

    def get_all_1st_segments_on_d2_section_order_by_skip_count_asc(self):
        def custom_condition(gdf):
            return (gdf[CABLE_LEVEL_FIELD_NAME] == 2) & (
                        gdf[CABLE_ORIGIN_FIELD_NAME] == gdf[CABLE_ORIGIN_BOX_FIELD_NAME])

        return self._dga.get_features_by_condition(custom_condition, sort_by=[CABLE_SKIP_COUNT_FIELD_NAME])

    def get_all_1st_segments_start_with_box_order_by_code_asc(self, box_code: str, upper_section: str):
        def custom_condition(gdf):
            return ((gdf[CABLE_ORIGIN_FIELD_NAME] == box_code)
                    & (gdf[CABLE_SECTION_FIELD_NAME] != upper_section))

        return self._dga.get_features_by_condition(custom_condition, sort_by=[CABLE_CODE_FIELD_NAME])

    def has_at_least_2_segments_on_cable(self, first_segment_data) -> bool:
        """
        是否有至少2段section在这条cable上
        section.ORIGIN==第一个section.EXTREMITY,
        section.SECTION = 第一个section.SECTION,
        :param first_segment_data: cable上的第一个section
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return (gdf[CABLE_ORIGIN_FIELD_NAME] == first_segment_data[CABLE_EXTREMITY_FIELD_NAME]) & (
                    gdf[CABLE_SECTION_FIELD_NAME] == first_segment_data[CABLE_SECTION_FIELD_NAME])

        return self._dga.get_count_by_condition(custom_condition) > 0

    def get_all_cables_start_with_one_point_by_orders(self, nap_code, sort_by: Optional[list[str]] = None,
                                                      ascending: bool | list[bool] = True):
        """
        获取指定点位为起点的所有线缆,根据字段排序
        :param nap_code: 根据nap_code查询
        :param sort_by: 排序字段
        :param ascending: 是否升序，默认True
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return gdf["origin_box"] == nap_code

        return self._dga.get_features_by_condition(condition=custom_condition, sort_by=sort_by, ascending=ascending)

    def get_all_1st_segments_start_with_one_point_by_orders(self, box_code, sort_by: Optional[list[str]] = None,
                                                            ascending: bool | list[bool] = True):
        """
        获取指定点位为起点的所有线缆,根据字段排序
        :param box_code: 根据nap_code查询
        :param sort_by: 排序字段
        :param ascending: 是否升序，默认True
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return (gdf[CABLE_ORIGIN_BOX_FIELD_NAME] == box_code) & (gdf[CABLE_ORIGIN_FIELD_NAME] == box_code)

        return self._dga.get_features_by_condition(condition=custom_condition, sort_by=sort_by, ascending=ascending)

    def get_all_cables_start_with_one_point_order_by_code_asc(self, nap_code):
        """
        获取指定点位为起点的所有线缆,根据code升序
        :param nap_code: 根据nap_code查询
        :return: 线缆列表
        """
        return self.get_all_cables_start_with_one_point_by_orders(nap_code=nap_code, sort_by=["code"], ascending=True)

    def get_all_cables_start_with_one_point(self, nap_code):
        """
        获取指定点位为起点的所有线缆
        :param nap_code: 根据nap_code查询
        :return: 线缆列表
        """
        return self.get_all_cables_start_with_one_point_by_orders(nap_code)

    def get_all_d2_cables_order_by_skip_count(self):
        """
        获取指定点位为起点的所有线缆
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return gdf[CABLE_LEVEL_FIELD_NAME] == 2

        return self._dga.get_features_by_condition(condition=custom_condition, sort_by=[CABLE_PORT_START_FIELD_NAME],
                                               ascending=True)

    def get_all_1st_segments_on_d3_cable_order_by_skip_count(self):
        """
        获取指定点位为起点的所有线缆
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return (gdf[CABLE_LEVEL_FIELD_NAME] == 3) & (
                        gdf[CABLE_ORIGIN_BOX_FIELD_NAME] == gdf[CABLE_ORIGIN_FIELD_NAME])

        return self._dga.get_features_by_condition(condition=custom_condition, sort_by=[CABLE_PORT_START_FIELD_NAME],
                                               ascending=True)

    def get_all_d3_cables_order_by_skip_count(self):
        """
        获取指定点位为起点的所有线缆
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return gdf[CABLE_LEVEL_FIELD_NAME] == 3

        return self._dga.get_features_by_condition(condition=custom_condition, sort_by=[CABLE_PORT_START_FIELD_NAME],
                                               ascending=True)

    def get_all_d2_d3_cables_order_by_skip_count(self):
        """
        获取指定点位为起点的所有线缆
        :return: 线缆列表
        """

        def custom_condition(gdf):
            return (gdf[CABLE_LEVEL_FIELD_NAME] == 2) | (gdf[CABLE_LEVEL_FIELD_NAME] == 3)

        return self._dga.get_features_by_condition(condition=custom_condition,
                                               sort_by=[CABLE_LEVEL_FIELD_NAME, CABLE_PORT_START_FIELD_NAME],
                                               ascending=True)

    def get_all_1st_segments_start_with_one_point(self, nap_code):
        """
        获取指定点位为起点的所有线缆
        :param nap_code: 根据nap_code查询
        :return: 线缆列表
        """
        return self.get_all_1st_segments_start_with_one_point_by_orders(nap_code)

    def set_extremity_by_cable_codes(self, code_extremity_dict):
        for cable_code, extremity in code_extremity_dict.items():
            def custom_condition(gdf):
                return gdf["code"] == cable_code

            self._dga.update_attributes(custom_condition, field="extremity", new_value=extremity)
        self._dga.save_changes(overwrite=True)

    def get_sub_cables_amt(self, nap_code):
        return self._dga.get_count_by_attribute("origin_box", "==", nap_code)

    def init_data_of_all_distribution01(self):
        """
        初始化所有sro节点的skip_count值为0，in_start为0
        """

        def custom_condition(gdf):
            return gdf["level"] == 1

        update_success_1 = self._dga.update_attributes(condition=custom_condition, field='skip_count', new_value=0)
        update_success_2 = self._dga.update_attributes(condition=custom_condition, field='port_start', new_value=1)
        if update_success_1 and update_success_2:
            self._dga.save_changes(overwrite=True)

    def update_skip_count_of_cable_start_with_point(self, start_point_code, start_point_skip_count):
        """
        更新起点为指定点的所有线缆的skip_count
        :param start_point_code: 根据nap_code查询
        :param start_point_skip_count: nap上的skip_count
        """

        def custom_condition(gdf):
            return gdf[CABLE_ORIGIN_BOX_FIELD_NAME] == start_point_code

        update_success = self._dga.update_attributes(condition=custom_condition, field='skip_count',
                                                 # new_value=self.gdf['port_start'] + start_point_skip_count - 1)
                                                 new_value=self.dga.gdf['port_start'] + start_point_skip_count - 1)
        if update_success:
            self._dga.save_changes(overwrite=True)

    def update_skip_count_of_1st_segment_of_section_start_with_point(self, start_point_code, start_point_skip_count):
        """
        更新起点为指定点的所有线缆的skip_count
        :param start_point_code: 根据nap_code查询
        :param start_point_skip_count: nap上的skip_count
        """

        def custom_condition(gdf):
            return (gdf[CABLE_ORIGIN_BOX_FIELD_NAME] == start_point_code) & (
                    gdf[CABLE_ORIGIN_FIELD_NAME] == start_point_code)

        update_success = self._dga.update_attributes(condition=custom_condition, field_values={
            CABLE_SKIP_COUNT_FIELD_NAME: self._dga.gdf[CABLE_PORT_START_FIELD_NAME] + start_point_skip_count - 1})
        if update_success:
            self._dga.save_changes(overwrite=True)
