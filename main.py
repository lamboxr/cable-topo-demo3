# This is a sample Python script.
import time

from box_sheet_creator import BoxSheetCreator
from data_service.service_manager import ServiceManager
# from data_service.service_manager import ServiceManager, init_global_services, global_services
from data_service.sro_service import SROService
from gen_topo_from_point import TopoFileGenerator
from init_data import DataInit


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def _main(params):
    # sro_gpkg_path = params['SRO']['gpkg_path']
    # sro_layer_name = params['SRO']['layer_name']
    # box_gpkg_path = params['BOX']['gpkg_path']
    # box_layer_name = params['BOX']['layer_name']
    # cable_gpkg_path = params['CABLE']['gpkg_path']
    # cable_layer_name = params['CABLE']['layer_name']
    # third_party_params = {
    #     "box_gpkg": box_gpkg_path,
    #     "cable_gpkg": cable_gpkg_path,
    #     "sro_gpkg": sro_gpkg_path,
    #     "box_layer": box_layer_name,
    #     "sro_layer": sro_layer_name,
    #     "cable_layer": cable_layer_name
    # }
    # from data_service.service_manager import init_global_services
    # init_global_services(**third_party_params)
    # 全局服务管理器实例（程序启动时自动初始化）

    services = ServiceManager(params)

    print(f"初始化后global_services: {services}")
    data_init = (
        DataInit(services))
    box_sheet_creator = BoxSheetCreator(services)
    topo_file_generator = TopoFileGenerator(services,box_sheet_creator)
    data_init.init_metadata()
    data_init.update_skip_count()
    return topo_file_generator.gen_topo_files('./')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
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
    print(_main(p))

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
