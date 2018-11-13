#!/usr/bin/python
# -*- coding:utf-8 -*-

import math
import numpy as np


if __name__ == "__main__":
    learning_rate = 0.01
    for a in np.arange(1, 10, 0.1):
        cur = 0
        for i in range(1000):
            cur -= learning_rate*(cur**2 - a)
        print ' %f的平方根(近似)为：%.8f，真实值是：%.8f' % (a, cur, math.sqrt(a))
