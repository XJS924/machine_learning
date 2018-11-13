#!/usr/bin/python
# -*- coding:utf-8 -*-

import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt


def clip(x, path):
    for i in range(len(x)):
        if x[i] >= path:
            x[i] %= path


if __name__ == "__main__":
    mpl.rcParams['font.sans-serif'] = [u'SimHei']
    mpl.rcParams['axes.unicode_minus'] = False

    path = 300     # 环形公路的长度
    n = 100         # 公路中初始车辆的数目
    v0 = 50         # 车辆的初始速度
    p = 0.3         # 随机减速概率
    Times = 200

    np.random.seed(0)
    x = np.random.rand(n) * path
    x.sort()    # 每辆车的位置
    v = np.tile([v0], n).astype(np.float)   # 每辆车的车速

    plt.figure(figsize=(10, 8), facecolor='w')
    np.set_printoptions(edgeitems=100, linewidth=3000)
    for t in range(Times):
        if x[0] > 50:
            v = np.concatenate(([v0], v))
            x = np.concatenate(([1], x))
        plt.scatter(x, [t]*len(x), s=5, c='k', alpha=0.5)
        n = len(v)  # 车辆数目有可能改变
        print x
        for i in range(n):
            d = (x[i+1] - x[i]) if i < n-1 else np.inf   # 距离前车的距离
            print d,
            if v[i] < d:    # 距离足够大
                if np.random.rand() > p:
                    v[i] += 1
                else:
                    v[i] -= 1
            else:
                v[i] = max(d-1, d*0.9)
        print
        v = v.clip(0, 150)
        x += v
        run_out = x < path
        x = x[run_out]
        v = v[run_out]
        # print t, 'x = ', x
        # print t, 'v = ', v
    plt.xlim(0, path)
    plt.ylim(0, Times)
    plt.xlabel(u'车辆位置', fontsize=16)
    plt.ylabel(u'模拟时间', fontsize=16)
    plt.title(u'环形公路车辆堵车模拟', fontsize=20)
    plt.tight_layout(pad=2)
    plt.show()
