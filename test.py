import geopandas as gpd

from constraints.field_name_mapper import BOX_CODE_FIELD_NAME
from data_service import data_service_cable, data_service_box, data_service_sro
from data_service.sro_service import SROService
from topo_generator import gen_topos


def join_query():
    # 1. 读取两个GeoPackage中的图层
    gpkg_cable = "cable.gpkg"
    gpkg_nap = "nap.gpkg"

    # 读取线图层（cable）和点图层（nap）
    gdf_cable = gpd.read_file(gpkg_cable, layer="cable")
    gdf_nap = gpd.read_file(gpkg_nap, layer="nap")

    # 2. 查看图层字段（确认关联字段存在）
    print("cable图层字段：", gdf_cable.columns.tolist())  # 需包含 "code" 字段
    print("nap图层字段：", gdf_nap.columns.tolist())  # 需包含 "cable_code" 字段

    # 3. 基于 "cable.code = nap.cable_code" 进行连接
    # left_on：左图层（cable）的关联字段
    # right_on：右图层（nap）的关联字段
    gdf_joined = gdf_cable.merge(
        gdf_nap,
        left_on="code",  # cable图层的关联字段
        right_on="cable_code",  # nap图层的关联字段
        how="inner"  # 连接方式（内连接，仅保留匹配的记录）
    )

    # 4. 查看连接结果
    print(f"连接后要素数量：{len(gdf_joined)}")
    # 显示关键关联字段，确认连接是否正确
    print(gdf_joined[["code", "cable_code"]].head())

def test_sorted_query():
    l = data_service_cable.get_all_cables_start_with_one_point_by_orders('CL11',['type','code'], ascending=False)
    if l is not None and not l.empty:
        for i, cable in l.iterrows():
            print(f"{cable['code']} {cable['type']} ")

    l2 = data_service_box.get_all_points_on_cable_by_order_in_start_asc('SRO01-1')
    if l2 is not None and not l2.empty:
        for j, nap in l2.iterrows():
            print(f"{nap['code']} {nap['in_start']}")

def test_count():
    amt = data_service_cable.get_sub_cables_amt("SRO001")
    print(amt)

def test_has_at_least_2_sections():
    sro = data_service_sro.get_sro_by_code("SRO-ELJ-QAE-0001")
    print(sro)
    print(f"=*10")
    print(type(sro['CODE']))
    _1st_section_list = data_service_cable.get_all_1st_segments_start_with_box_order_by_code_asc(sro[BOX_CODE_FIELD_NAME])
    for _, section in _1st_section_list.iterrows():
        bl = data_service_cable.has_at_least_2_segments_on_cable(section)
        print(bl)

def test1():
    sro_service = SROService(gpkg_path='./gpkg/SRO.gpkg', layer_name='elj_qae_sro')
    a =sro_service.get_all_sro_order_by_code_asc()
    print(a)

def main_test():
    p = {
        "SRO": {
            "gpkg_path": "./gpkg/SRO.gpkg",
            "layer_name": "elj_qae_sro"
        },
        "BOX": {
            "gpkg_path": "./gpkg/BOX.gpkg",
            "layer_name": "elj_qae_boite_optique"
        },
        "CABLE": {
            "gpkg_path": "./gpkg/CABLE.gpkg",
            "layer_name": "elj_qae_cable_optique"
        }
    }
    print(gen_topos(p, './'))

if __name__ == '__main__':
    main_test()
