# This is a sample Python script.
import os
import pathlib
import time
from datetime import datetime

from data_service.service_manager import ServiceManager
# from data_service.service_manager import ServiceManager, init_global_services, global_services
from data_service.sro_service import SROService
from topo_creator.box_sheet_creator import BoxSheetCreator
from topo_creator.gen_topo_from_point import TopoFileGenerator
from topo_creator.init_data import DataInit
from utils.zip_utils import zip_files


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def gen_topos(params, output_dir):
    try:
        services = ServiceManager(params)

        print(f"初始化后global_services: {services}")
        data_init = (
            DataInit(services))
        box_sheet_creator = BoxSheetCreator(services)
        topo_file_generator = TopoFileGenerator(services, box_sheet_creator)
        data_init.init_metadata()
        data_init.update_skip_count()
        pathlib.Path(output_dir).mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        file_list = topo_file_generator.gen_topo_files(output_dir, timestamp)
        if not file_list or len(file_list):

            dl_file = os.path.join(output_dir, f"SRO_TOPO_{timestamp}.zip")
            zip_files(file_list, dl_file)
            return {
                "code": 200,
                "file_path": dl_file
            }
        else:
            return {
                "code": 400,
                "file_path": None,
                "error_message": f"没有找到SRO信息"
            }
    except Exception as e:
        return {
            "code": 500,
            "file_path": None,
            "error_message": f"发生未知错误：{e}"
        }


# Press the green button in the gutter to run the script.


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
