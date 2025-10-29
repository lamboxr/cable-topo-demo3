# data_service/sro_service.py
from typing import List

from constraints.field_name_mapper import BOX_CODE_FIELD_NAME, BOX_SKIP_COUNT_FIELD_NAME
from data_service.base import BaseLayerService


class SROService(BaseLayerService):
    def __init__(self, gpkg_path, layer_name):
        # 初始化父类：指定SRO对应的GPKG和图层名
        super().__init__(gpkg_path=gpkg_path, layer_name=layer_name)

    def get_all_sro_order_by_code_asc(self) -> List[str]:
        return self._dga.get_features_by_condition(sort_by=[BOX_CODE_FIELD_NAME])

    def get_sro_by_code(self, code):
        return self._dga.get_features_by_attribute(field=BOX_CODE_FIELD_NAME, op="==", value=code)

    def init_data_of_all_sro_points(self):
        """
        初始化所有sro节点的skip_count值为0
        """

        # def custom_condition(gdf):
        #     return gdf["class"] == 'SRO'
        update_success = self._dga.update_attributes(condition=None, field_values={BOX_SKIP_COUNT_FIELD_NAME: 0})
        if update_success:
            self._dga.save_changes(overwrite=True)
