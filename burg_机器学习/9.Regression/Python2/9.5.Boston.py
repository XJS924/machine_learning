#!/usr/bin/python
# -*- coding:utf-8 -*-

import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.linear_model import ElasticNetCV
import sklearn.datasets
from pprint import pprint
from sklearn.preprocessing import PolynomialFeatures, StandardScaler, MinMaxScaler
from sklearn.pipeline import Pipeline
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.ensemble import RandomForestRegressor
import warnings
import exceptions


def not_empty(s):
    return s != ''


def log_scaler(data):
    for i in data.columns[:-1]:
        data[i] = np.log(data[i] + 1)
    return data


if __name__ == "__main__":
    warnings.filterwarnings(action='ignore')
    pd.set_option('display.width', 200)
    data = pd.read_fwf('housing.data', header=None)
    # data = sklearn.datasets.load_boston()
    # x = np.array(data.data)
    # y = np.array(data.target)
    x = MinMaxScaler().fit_transform(data.iloc[:, :-1])
    for i in data.columns[:-1]:
        x[i] = np.log(x[i] - x[i].min() + 1)
    y = data.iloc[:, -1]
    print u'样本个数：%d, 特征个数：%d' % x.shape

    x_train, x_test, y_train, y_test = train_test_split(x, y, train_size=0.7, random_state=0)
    model = Pipeline([
        ('poly', PolynomialFeatures(degree=2, include_bias=True, interaction_only=False)),
        ('linear', ElasticNetCV(l1_ratio=[0.1, 0.3, 0.5, 0.7, 0.99, 1], alphas=np.logspace(-3, 2, 5),
                                fit_intercept=False, max_iter=1e3, cv=3))
    ])
    print u'开始建模...'
    model.fit(x_train, y_train)
    linear = model.get_params('linear')['linear']
    print u'超参数：', linear.alpha_
    print u'L1 ratio：', linear.l1_ratio_
    print u'系数：', linear.coef_.ravel()

    order = y_test.argsort()
    y_test = y_test.values[order.values]
    x_test = x_test[order]
    y_train_pred = model.predict(x_train)
    print u'训练集R2:', r2_score(y_train, y_train_pred)
    print u'训练集均方误差：', mean_squared_error(y_train, y_train_pred)
    y_test_pred = model.predict(x_test)
    print u'测试集R2:', r2_score(y_test, y_test_pred)
    print u'测试集均方误差：', mean_squared_error(y_test, y_test_pred)

    t = np.arange(len(y_test_pred))
    mpl.rcParams['font.sans-serif'] = [u'simHei']
    mpl.rcParams['axes.unicode_minus'] = False
    plt.figure(facecolor='w')
    plt.plot(t, y_test, 'r-', lw=2, label=u'真实值')
    plt.plot(t, y_test_pred, 'g-', lw=2, label=u'估计值')
    plt.legend(loc='upper left')
    plt.title(u'波士顿房价预测', fontsize=18)
    plt.xlabel(u'样本编号', fontsize=15)
    plt.ylabel(u'房屋价格', fontsize=15)
    plt.grid()
    plt.show()
