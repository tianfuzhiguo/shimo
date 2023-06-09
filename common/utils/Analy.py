# -*- coding: utf-8 -*-
from __future__ import print_function

class Analy:
    # 定义generator函数，参数indict为需要解析的字典或列表，pre为键或索引列表
    def generator(self, indict, pre=None):
        # 判断pre是否为空，是则将其转化为空列表，否则复制一份到新的列表pre中
        pre = pre[:] if pre else []
        if isinstance(indict, dict): # 如果indict是字典类型
            for key, value in indict.items():
                if isinstance(value, dict): # 如果value是字典
                    if not value: # 如果value为空字典，则生成由pre、key和空字典组成的元组
                        yield pre + [f"['{key}']", {}]
                    else: # 如果value不为空字典，则递归调用自身生成元组
                        for d in self.generator(value, pre + [f"['{key}']"]):
                            yield d
                elif isinstance(value, list): # 如果value是列表
                    if not value: # 如果value为空列表，则生成由pre、key和空列表组成的元组
                        yield pre + [f"['{key}']", []]
                    else: # 如果value不为空列表，则遍历value中每一个元素并根据类型生成相应元组
                        for i, v in enumerate(value):
                            if isinstance(v, dict):
                                for d in self.generator(v, pre + [f"['{key}'][{i}]"]):
                                    yield d
                            else:
                                yield pre + [f"['{key}'][{i}]", v]

                elif isinstance(value, tuple): # 如果value是元组
                    if not value: # 如果value为空元组，则生成由pre、key和空元组组成的元组
                        yield pre + [f"['{key}']", ()]
                    else: # 如果value不为空元组，则遍历value中每一个元素并根据类型生成相应元组
                        for i, v in enumerate(value):
                            for d in self.generator(v, pre + [f"['{key}'][{i}]"]):
                                yield d
                else: # 如果value是其他类型
                    yield pre + [f"['{key}']", value] # 则生成由pre、key和value组成的元组
        elif isinstance(indict, list): # 如果indict是列表类型
            if not indict: # 如果indict为空列表，则生成由pre和空列表组成的元组
                yield pre + [f"[{0}]", []]
            else: # 如果indict不为空列表，则遍历indict中每一个元素并根据类型生成相应元组
                for i, v in enumerate(indict):
                    if isinstance(v, dict):
                        for d in self.generator(v, pre + [f"[{i}]"]):
                            yield d
                    else:
                        yield pre + [f"[{i}]", v]
        else: # 如果indict是其他类型
            yield [indict] # 则生成由indict组成的元组

    # 定义analy函数，参数js为需要解析的字典或列表
    def analy(self, js):
        key = [] # 定义存储键的列表
        value = [] # 定义存储值的列表
        for item in self.generator(js): # 遍历generator生成的元组
            k = ''.join(item[:-1])
            v = item[-1]
            key.append(k) # 将键添加到key列表中
            value.append(v) # 将值添加到value列表中
        return key, value # 返回键和值的列表
