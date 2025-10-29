# 获取所有distribution1
from constraints.field_name_mapper import CABLE_CODE_FIELD_NAME, CABLE_SKIP_COUNT_FIELD_NAME, CABLE_SECTION_FIELD_NAME, \
    BOX_CODE_FIELD_NAME, BOX_SKIP_COUNT_FIELD_NAME

global_services = None
class DataInit:
    def __init__(self, data_services) -> None:
        # 仅初始化一次（使用类级别标记判断）
        # 初始化三个服务（通过third_party_params获取路径和图层名）
        self.sro_service = data_services.sro_service
        self.box_service = data_services.box_service
        self.cable_service = data_services.cable_service

    def fill_extremity_of_all_cables(self):
        extremities = self.box_service.get_all_extremities()
        if extremities is not None and not extremities.empty:
            _dict = {}
            for idx, _nap in extremities.iterrows():
                _dict[_nap["cable_in"]] = _nap["code"]
                print(f'======={_nap["code"]}')
            self.cable_service.set_extremity_by_cable_codes(_dict)



    def init_metadata(self):
        # 所有SRO的点skip_count都设置为0
        self.sro_service.init_data_of_all_sro_points()

        # 所有distribution01类型的skip_count都设置为0
        # self.cable_service.init_data_of_all_distribution01()

        # fill_extremity_of_all_cables()

    def update_skip_count_start_with_one_point(self, box_code, nap_skip_count):
        # 更新从单个点（起点、掏芯点、终点）上分离出去的所有子线缆上的skip count
        self.cable_service.update_skip_count_of_1st_segment_of_section_start_with_point(box_code, nap_skip_count)
        # 获取单个点（起点、掏芯点、终点）上分离出去的所有子线缆
        _1st_segment_list = self.cable_service.get_all_1st_segments_start_with_one_point(box_code)
        if _1st_segment_list is not None and not _1st_segment_list.empty:
            for _segment_idx, _1st_segment in _1st_segment_list.iterrows():
                _cable_code = _1st_segment[CABLE_CODE_FIELD_NAME]
                _section = _1st_segment[CABLE_SECTION_FIELD_NAME]
                _cable_skip_count = _1st_segment[CABLE_SKIP_COUNT_FIELD_NAME]
                # 更在单个SECTION（d1,d2,d3）上掏芯的所有点（closure,终点）的skip count
                self.box_service.update_skip_count_of_boxs_on_section(_section, _cable_skip_count)
                # 获取在单个SECTION（d1,d2,d3）上掏芯的所有点（closure,终点）
                box_list = self.box_service.get_all_boxs_on_section(_section)
                if box_list is not None and not box_list.empty:
                    for _box_idx, _box in box_list.iterrows():
                        _box_code = _box[BOX_CODE_FIELD_NAME]
                        _box_skip_count = _box[BOX_SKIP_COUNT_FIELD_NAME]
                        self.update_skip_count_start_with_one_point(_box_code, _box_skip_count)


    def update_skip_count(self):
        all_sro_point = self.sro_service.get_all_sro_order_by_code_asc()
        if all_sro_point is not None and not all_sro_point.empty:
            for _sro_idx, _sro in all_sro_point.iterrows():
                _nap_code = _sro[BOX_CODE_FIELD_NAME]
                _nap_skip_count = _sro[BOX_SKIP_COUNT_FIELD_NAME]
                self.update_skip_count_start_with_one_point(_nap_code, _nap_skip_count)
    """==================主流程=================="""

    def init(self):
        self.init_metadata()
        self.update_skip_count()

# 更新所有distribution1线缆上的掏芯点上的skip_count值
# if __name__ == '__main__':
#     main()
