from typing import Optional

from utils import gpkg_utils
from utils.gda_utils import LayerDGA


def _gda():
    gpkg_cable_path = gpkg_utils.get_gpkg_path("cable.gpkg")
    layer_name = "cable"
    cable_dga = LayerDGA(gpkg_cable_path, layer_name)
    return cable_dga


__gda = _gda()


def get_all_cables_start_with_one_point_by_orders(nap_code, sort_by: Optional[list[str]] = None,
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

    return __gda.get_features_by_condition(condition=custom_condition, sort_by=sort_by, ascending=ascending)


def get_all_cables_start_with_one_point_order_by_code_asc(nap_code):
    """
    获取指定点位为起点的所有线缆,根据code升序
    :param nap_code: 根据nap_code查询
    :return: 线缆列表
    """
    return get_all_cables_start_with_one_point_by_orders(nap_code=nap_code, sort_by=["code"], ascending=True)


def get_all_cables_start_with_one_point(nap_code):
    """
    获取指定点位为起点的所有线缆
    :param nap_code: 根据nap_code查询
    :return: 线缆列表
    """
    return get_all_cables_start_with_one_point_by_orders(nap_code)


def set_extremity_by_cable_codes(code_extremity_dict):
    for cable_code, extremity in code_extremity_dict.items():
        def custom_condition(gdf):
            return gdf["code"] == cable_code

        __gda.update_attributes(custom_condition, field="extremity", new_value=extremity)
    __gda.save_changes(overwrite=True)

def get_sub_cables_amt(nap_code):
    return __gda.get_count_by_attribute("origin_box", "==", nap_code)


def init_data_of_all_distribution01():
    """
    初始化所有sro节点的skip_count值为0，in_start为0
    """

    def custom_condition(gdf):
        return gdf["level"] == 1

    update_success_1 = __gda.update_attributes(condition=custom_condition, field='skip_count', new_value=0)
    update_success_2 = __gda.update_attributes(condition=custom_condition, field='port_start', new_value=1)
    if update_success_1 and update_success_2:
        __gda.save_changes(overwrite=True)


def update_skip_count_of_cable_start_with_point(start_point_code, start_point_skip_count):
    """
    更新起点为指定点的所有线缆的skip_count
    :param start_point_code: 根据nap_code查询
    :param start_point_skip_count: nap上的skip_count
    """

    def custom_condition(gdf):
        return gdf["origin_box"] == start_point_code

    update_success = __gda.update_attributes(condition=custom_condition, field='skip_count',
                                             new_value=__gda.gdf['port_start'] + start_point_skip_count - 1)
    if update_success:
        __gda.save_changes(overwrite=True)


if __name__ == '__main__':
    gpkg_cable_path = "../gpkg/cable.gpkg"
    gda = LayerDGA(gpkg_cable_path, "cable")
    print(gda.sub_cables_amt('SRO001'))
