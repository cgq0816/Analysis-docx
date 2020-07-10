# -*- coding: utf-8 -*-
from models.model1 import Model1
from models.model2 import Model2
import logging


class Ensemble_Model(object):

    def __init__(self):

        self.model1 = Model1()
        self.model2 = Model2()

    def analysis(self, content):

        # 依赖 model1、model2 输出
        model1_output = self.model1.inference(content)
        model2_output = self.model2.inference(content)

        final_output = self.inference(model1_output, model2_output)

        return final_output

    def inference(self, x1, x2):

        y = x1 + x2 + '<FINAL_INFERENCE>'

        return y