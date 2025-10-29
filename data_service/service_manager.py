# data_service/service_manager.py
from typing import Optional

from data_service.box_service import BOXService
from data_service.cable_service import CABLEService
from data_service.sro_service import SROService



class ServiceManager:
    """服务管理器（单例），统一管理三个图层服务"""
    _instance: Optional["ServiceManager"] = None
    _initialized: bool = False  # 类级别的初始化标记（关键修正）

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, third_party_params) -> None:
        # 仅初始化一次（使用类级别标记判断）
        if not self._initialized:  # 现在可正常访问 _initialized
            # 初始化三个服务（通过third_party_params获取路径和图层名）
            sro_config = third_party_params["SRO"]
            box_config = third_party_params["BOX"]
            cable_config = third_party_params["CABLE"]

            self.sro_service = SROService(
                gpkg_path=sro_config["gpkg_path"],
                layer_name=sro_config["layer_name"]
            )
            self.box_service = BOXService(
                gpkg_path=box_config["gpkg_path"],
                layer_name=box_config["layer_name"]
            )
            self.cable_service = CABLEService(
                gpkg_path=cable_config["gpkg_path"],
                layer_name=cable_config["layer_name"]
            )
            self._initialized = True  # 标记为已初始化

# global_services: Optional["ServiceManager"] = None

# data_service/service_manager.py（续）
# def init_global_services(params) -> None:
#     global_services  # 显式声明为全局变量（关键）
#     if global_services is None:
#         global_services = ServiceManager(params)
#         print("全局服务初始化成功")
#     else:
#         print("全局服务已初始化")


