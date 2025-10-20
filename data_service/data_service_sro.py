import os
from typing import Optional

import pandas as pd

from constraints.field_name_mapper import BOX_CODE_FIELD_NAME
from data_service import data_service_cable
from utils import gpkg_utils
from utils.gda_utils import LayerDGA


def _gda():
    gpkg_cable_path = gpkg_utils.get_gpkg_path("SRO.gpkg")
    layer_name = "elj_qae_sro"
    nap_gda = LayerDGA(gpkg_cable_path, layer_name)
    return nap_gda


__gda = _gda()

def get_all_sro_order_by_code_asc():
    return __gda.get_features_by_condition(sort_by=[BOX_CODE_FIELD_NAME])


def get_sro_by_code(code):
    return __gda.get_features_by_attribute(field='CODE', op="==",value=code)