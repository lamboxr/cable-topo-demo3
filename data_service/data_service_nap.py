from typing import Optional

import pandas as pd

from data_service import data_service_cable
from utils import gpkg_utils
from utils.gda_utils import LayerDGA


def _gda():
    gpkg_cable_path = gpkg_utils.get_gpkg_path("nap.gpkg")
    layer_name = "nap"
    nap_gda = LayerDGA(gpkg_cable_path, layer_name)
    return nap_gda


__gda = _gda()


def get_all_extremities():
    """
    获取所有终点类型的节点
    :return: 终点类型的节点列表
    """
    return __gda.get_features_by_attribute('pass_seq', '==', 100)


def boxs_amt_on_cable(cable_box):
    return __gda.get_count_by_attribute("cable_in", "==", cable_box)


def get_all_sro_points_by_orders(sort_by=None, ascending=True):
    """
    获取所有SRO节点，根据字段排序
    :param sort_by: 排序字段
    :param ascending: 是否升序 默认true
    :return: 所有SRO节点列表
    """
    return __gda.get_features_by_attribute(field='class', op='==', value='SRO', sort_by=sort_by, ascending=ascending)


def get_all_sro_points_by_order_code_asc():
    """
    获取所有SRO节点，根据code升序
    :return: 所有SRO节点列表
    """
    return get_all_sro_points_by_orders(sort_by="code", ascending=True)


def get_all_sro_points():
    """
    获取所有SRO节点
    :return: 所有SRO节点列表
    """
    return get_all_sro_points_by_orders()


def get_all_points_on_cable_by_orders(cable_code: str, sort_by: Optional[list[str]] = None,
                                      ascending: bool | list[bool] = True):
    """
    获取指定线缆上所有掏芯节点（closure,终点）
    :param cable_code: 根据线缆code查询
    :param sort_by: 排序字段
    :param ascending: 是否升序 默认true
    :return: 所有掏芯节点列表
    """
    return __gda.get_features_by_attribute(field='cable_in', op='==', value=cable_code, sort_by=sort_by,
                                           ascending=ascending)


def get_all_points_on_cable_by_order_in_start_asc(cable_code: str):
    """
    获取指定线缆上所有掏芯节点（closure,终点）
    :param cable_code: 根据线缆code查询
    :return: 所有掏芯节点列表
    """
    return get_all_points_on_cable_by_orders(cable_code=cable_code, sort_by=['in_start'], ascending=True)


def get_all_points_on_cable(cable_code: str):
    """
    获取指定线缆上所有掏芯节点（closure,终点）
    :param cable_code: 根据线缆code查询
    :return: 所有掏芯节点列表
    """
    return get_all_points_on_cable_by_orders(cable_code=cable_code)


def init_data_of_all_sro_points():
    """
    初始化所有sro节点的skip_count值为0
    """

    def custom_condition(gdf):
        return gdf["class"] == 'SRO'

    update_success = __gda.update_attributes(condition=custom_condition, field='skip_count', new_value=0)
    if update_success:
        __gda.save_changes(overwrite=True)


def update_skip_count_of_points_on_cable(cable_code: str, cable_skip_count: int):
    """
    更新指定线缆上所有掏芯节点的skip_count
    :param cable_code: 根据线缆code查询
    :param cable_skip_count: 线缆skip_count
    """
    print(f"cable_code: {cable_code}, skip count: {cable_skip_count}")

    def custom_condition(gdf):
        return gdf["cable_in"] == cable_code

    __gda.update_attributes(condition=custom_condition, field="skip_count",
                            new_value=__gda.gdf['in_start'] + cable_skip_count - 1)
    __gda.save_changes(overwrite=True)
