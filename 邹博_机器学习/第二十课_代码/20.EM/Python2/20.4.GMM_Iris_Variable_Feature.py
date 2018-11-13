# -*- coding:utf-8 -*-

import numpy as np
import pandas as pd
from sklearn.mixture import GaussianMixture
from sklearn.metrics.pairwise import pairwise_distances_argmin


# 选取2 3 4 个特征分别用GMM训练，结果显示[0,2,3]特征能达到98%准确率，麻烦老师帮忙看下有什么问题，还有没有什么方法能提高准确率，多谢老师
if __name__ == '__main__':
    path = '..\\9.Regression\\iris.data'
    data = pd.read_csv(path,header=None)
    x_prime = data[np.arange(4)]
    y = pd.Categorical(data[4]).codes

    #特征选择
    feature_pairs = [[0, 1], [0, 2], [0, 3], [1, 2], [1, 3], [2, 3],[0,1,2],[0,1,3],[0,2,3],[1,2,3],[0,1,2,3]]
    print(feature_pairs)
    n_component = 3
    for k, pair in enumerate(feature_pairs, start= 1):
        #4 feature
        print('第%d个特征组合进行训练 ' %k ,pair)
        x = x_prime[pair]
        print(np.shape(x))
        gmm = GaussianMixture(n_components=n_component,covariance_type='full',random_state=0)
        gmm.fit(x)

        # print('预测均值 = \n', gmm.means_)
        # print('预测方差 = \n', gmm.covariances_)

        y_hat = gmm.predict(x)
        m = np.array([np.mean(x[y == i], axis= 0) for i in range(3)])
        # print(m)
        order = pairwise_distances_argmin(m, gmm.means_, axis= 1,metric='euclidean')
        print("顺序 \t", order)

        #重新映射
        n_sample = y.size
        # print(n_sample)
        n_types = 3
        change = np.empty((n_types, n_sample), dtype= np.bool)
        for i in range(n_types):
            change[i] = y_hat == order[i]
        for i in range(n_types):
            y_hat[change[i]] = i

        acc = '准确率%.2f%%'% (100*np.mean(y_hat == y))
        print(acc)


