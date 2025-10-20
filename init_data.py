# 获取所有distribution1
from data_service import data_service_cable
from data_service import data_service_box

def fill_extremity_of_all_cables():
    extremities = data_service_nap.get_all_extremities()
    if extremities is not None and not extremities.empty:
        _dict = {}
        for idx, _nap in extremities.iterrows():
            _dict[_nap["cable_in"]] = _nap["code"]
            print(f'======={_nap["code"]}')
        data_service_cable.set_extremity_by_cable_codes(_dict)



def init_metadata():
    # 所有SRO的点skip_count都设置为0
    data_service_nap.init_data_of_all_sro_points()

    # 所有distribution01类型的skip_count都设置为0
    data_service_cable.init_data_of_all_distribution01()

    fill_extremity_of_all_cables()

def update_skip_count_start_with_one_point(nap_code, nap_skip_count):
    # 更新从单个点（起点、掏芯点、终点）上分离出去的所有子线缆上的skip count
    data_service_cable.update_skip_count_of_cable_start_with_point(nap_code, nap_skip_count)
    # 获取单个点（起点、掏芯点、终点）上分离出去的所有子线缆
    cable_list = data_service_cable.get_all_cables_start_with_one_point(nap_code)
    if cable_list is not None and not cable_list.empty:
        for _cable_idx, _cable in cable_list.iterrows():
            _cable_code = _cable["code"]
            _cable_skip_count = _cable["skip_count"]
            # 更在单个线缆（d1,d2,d3）上掏芯的所有点（closure,终点）的skip count
            data_service_nap.update_skip_count_of_points_on_cable(_cable_code, _cable_skip_count)
            # 获取在单个线缆（d1,d2,d3）上掏芯的所有点（closure,终点）
            nap_list = data_service_nap.get_all_points_on_cable(_cable_code)
            if nap_list is not None and not nap_list.empty:
                for _nap_idx, _nap in nap_list.iterrows():
                    _nap_code = _nap["code"]
                    _nap_skip_count = _nap["skip_count"]
                    update_skip_count_start_with_one_point(_nap_code, _nap_skip_count)


def update_skip_count():
    all_sro_point = data_service_nap.get_all_sro_points()
    if all_sro_point is not None and not all_sro_point.empty:
        for _sro_idx, _sro in all_sro_point.iterrows():
            _nap_code = _sro["code"]
            _nap_skip_count = _sro["skip_count"]
            update_skip_count_start_with_one_point(_nap_code, _nap_skip_count)
"""==================主流程=================="""

def main():
    init_metadata()
    update_skip_count()

# 更新所有distribution1线缆上的掏芯点上的skip_count值
if __name__ == '__main__':
    main()
