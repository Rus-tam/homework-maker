# -*- coding: utf-8 -*-
"""
Created on Wed Sep 21 21:30:07 2022

@author: Rustam
"""

from connector import Connector
from writer import Writer

class StreamReader:
    def __init__(self, case):
        self.case = case
        self.streams_root = self.case.Flowsheet.MaterialStreams
        self.operations_root = self.case.Flowsheet.Operations
    
    def get_all_exchangers_names(self):
        operations_root = self.operations_root
        exchangers = []
        for name in operations_root.Names:
            if operations_root.Item(name).TypeName == "heatexop" or operations_root.Item(name).TypeName == "coolerop" or operations_root.Item(name).TypeName == "heaterop":
                exchangers.append(operations_root.Item(name))
        return exchangers
    
    def get_streams_info(self):
        stream_root = self.streams_root
        stream_names = self.streams_root.Names
        streams_info = []
        for name in stream_names:
            stream_data = {"Name": str(name), "Temperature, C": int(stream_root.Item(name).Temperature), 
                           "Mass flow, kg/h": int(stream_root.Item(name).MassFlowValue * 3600), 
                           "Mass Heat Capacity, kJ/kg*C": int(stream_root.Item(name).MassHeatCapacityValue)}
            streams_info.append(stream_data)
        return streams_info
    
    def get_exchangers_info(self):
        exchangers = self.get_all_exchangers_names()
        duties = []
        for item in exchangers:
            duty = {"Name": item.name, "Duty, kJ/h": int(item.DutyValue * 3600)}
            duties.append(duty)
        
        return duties
        
        
            

    






file_path = r"C:\Users\Rustam\Documents\stream_reader\exchanger-1.hsc"
case = Connector.case(file_path)

stream_reader = StreamReader(case)
duties = stream_reader.get_exchangers_info()
stream_info = stream_reader.get_streams_info()
# print(duties)
# print('++++++++++')
# print(stream_info)
Writer.make_excel(duties, stream_info)
