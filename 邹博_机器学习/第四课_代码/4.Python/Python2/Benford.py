# -*- coding:utf-8 -*-
# /usr/bin/python

import numpy as np


def top(number):
    number -= int(number)
    return int(10 ** number)

if __name__ == '__main__':
    x = np.arange(1, 1001)
    y = np.cumsum(np.log10(x))
    bf = np.zeros(9, dtype=np.int)
    for t in y:
        bf[top(t) - 1] += 1
    print bf
