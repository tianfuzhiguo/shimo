# -*- coding: utf-8 -*-
from __future__ import print_function

'''
@author: dujianxiao
'''
class analy():
    def generator(self,indict, pre=None):
        pre = pre[:] if pre else []
        if isinstance(indict, dict):
            for key, value in indict.items():
                if isinstance(value, dict):
                    if len(value) == 0:
                        yield pre+["['"+key+"']", '{}']
                    else:
                        for d in self.generator(value, pre + ["['"+key+"']"]):
                            yield d
                elif isinstance(value, list):
                    if len(value) == 0:    
                        yield pre+["['"+key+"']", '[]']#
                    else:
                        for v in value:
                            if isinstance(v, dict):
                                for d in self.generator(v, pre + ["['"+key+"']"+str([value.index(v)])]):
                                    yield d
                            else:
                                yield pre + ["['"+key+"']"+str([value.index(v)]),v]
                    
                elif isinstance(value, tuple):
                    if len(value) == 0:
                        yield pre+["['"+key+"']", '()']
                    else:
                        for v in value:
                            for d in self.generator(v, pre + ["['"+key+"']"]):
                                yield d
                else:
                    yield pre + ["['"+key+"']", value]
        elif isinstance(indict, list):
            if len(indict) == 0:    
                yield pre+[str([0]), '[]']
            else:
                for v in indict:      
                    if isinstance(v, dict):
                        for d in self.generator(v, pre + [str([indict.index(v)])]):
                            yield d
                    else:
                        yield pre + [str([indict.index(v)]),v]
        else:
            yield indict
    
    def analy(self,js):
        key=[]
        value=[]
        for item in self.generator(js):
            ss=(''.join(item[0:-1])), item[-1]
            ss=list(ss)
            key.append(ss[0])
            value.append(ss[-1])
        return key,value
    
    