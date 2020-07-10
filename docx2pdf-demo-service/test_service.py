# -*- coding: utf-8 -*-

import io
import os
import requests
import json
import random
from functools import wraps
import time


def timefn(fn):
    """计算函数耗时的修饰器"""
    @wraps(fn)
    def measure_time(*args, **kwargs):
        t1 = time.time()
        result = fn(*args, **kwargs)
        t2 = time.time()
        print("@timefn: " + fn.__name__ + " took: " + str(t2 - t1) + " seconds")
        return result
    return measure_time


@timefn
def server_test(url, post_data):

    url = 'http://10.0.0.95:3101/analysis'
    response = requests.post(url, data=json.dumps(post_data))
    print('status: ', response)
    response = response.json()
    print('response: ', response)


if __name__ == '__main__':

    url = 'http://localhost:31001'

    doc_url = 'https://rxhui-crawl.oss-cn-beijing.aliyuncs.com/fileupload/2020063010/71efcbb9b4ce4f948884e04db9b479ec.docx'
    # doc_url = "http://10.0.0.112/group112/M00/0D/35/CgAAb17pxNSLT26oAAGN3Z3P4vw22.docx"

    post_data = {'doc_url': doc_url
                 }

    server_test(url, post_data)




